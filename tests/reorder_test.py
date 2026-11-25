import bisect
import itertools
import unittest
from typing import Any, Dict, List, Optional, Sequence
from unittest.mock import Mock

from gspread.exceptions import (
    DuplicateIDsError,
    IDSetMismatchError,
    InvalidInputValue,
    MissingIDsError,
)
from gspread.http_client import HTTPClient
from gspread.reorder import RowReorderPlan, plan_row_reorder
from gspread.utils import ValueRenderOption
from gspread.worksheet import Worksheet

TITLE = "Reconcile"
SHEET_ID = 42
SPREADSHEET_ID = "spreadsheet-id"


def lis_length(sequence: Sequence[int]) -> int:
    tails: List[int] = []
    for value in sequence:
        position = bisect.bisect_left(tails, value)
        if position == len(tails):
            tails.append(value)
        else:
            tails[position] = value
    return len(tails)


class FakeGrid:
    """In-memory sheet state and the two client calls the planner uses.

    ``batch_update`` applies ``moveDimension`` requests with the real Sheets
    semantics: ``destinationIndex`` is interpreted against the coordinates
    from before the source row is removed.
    """

    def __init__(self, rows: List[List[Any]], first_row: int):
        self.rows = [list(row) for row in rows]
        self.first_row = first_row
        self.batch_calls: List[Dict[str, Any]] = []
        self.value_gets: List[Dict[str, Any]] = []

    # -- reads ---------------------------------------------------------
    def values_get(
        self, spreadsheet_id: str, range: str, params=None
    ) -> Dict[str, Any]:
        assert spreadsheet_id == SPREADSHEET_ID
        self.value_gets.append({"range": range, "params": dict(params or {})})
        a1 = range.split("!")[1]
        start = a1.split(":")[0]
        letters = "".join(ch for ch in start if ch.isalpha())
        column = 0
        for ch in letters:
            column = column * 26 + (ord(ch.upper()) - 64)
        column -= 1

        start_row = int("".join(ch for ch in start if ch.isdigit()))
        end = a1.split(":")[-1]
        end_digits = "".join(ch for ch in end if ch.isdigit())
        end_row = int(end_digits) if end_digits else None
        offset = start_row - self.first_row
        values = [row[column] if column < len(row) else "" for row in self.rows]
        values = values[offset:]
        if end_row is not None:
            values = values[: end_row - start_row + 1]
        # Sheets omits trailing blanks from the payload.
        while values and values[-1] == "":
            values.pop()
        return {"values": [values] if values else []}

    # -- writes --------------------------------------------------------
    def batch_update(self, spreadsheet_id: str, body: Dict[str, Any]) -> Dict[str, Any]:
        assert spreadsheet_id == SPREADSHEET_ID
        self.batch_calls.append(body)
        for request in body["requests"]:
            move = request["moveDimension"]
            source = move["source"]
            assert source["sheetId"] == SHEET_ID
            assert source["dimension"] == "ROWS"
            assert source["endIndex"] - source["startIndex"] == 1
            # Sheet indexes are zero-based; the first data row sits at
            # zero-based index ``first_row - 1``.
            start = source["startIndex"] - (self.first_row - 1)
            destination = move["destinationIndex"] - (self.first_row - 1)
            row = self.rows.pop(start)
            # destinationIndex refers to pre-removal coordinates
            effective = destination - 1 if destination > start else destination
            self.rows.insert(effective, row)
        return {"replies": [{} for _ in body["requests"]]}

    # -- helpers -------------------------------------------------------
    def id_column(self, column: int = 0) -> List[Any]:
        return [row[column] for row in self.rows]


def make_worksheet(
    ids: Sequence[Any], first_row: int = 2, value_column: Optional[List[Any]] = None
) -> tuple:
    """Return ``(worksheet, grid)`` with one data row per ID.

    The ID sits in column A; column B carries a per-row payload so tests can
    check that whole rows travel with their ID.
    """
    rows = [[id_, "payload-{}".format(id_)] for id_ in ids]
    grid = FakeGrid(rows, first_row)
    if value_column is not None:
        grid.rows = [list(row) for row in value_column]
    # A real HTTPClient instance is required by the Worksheet constructor;
    # its network calls are never reached because the methods are replaced.
    client = HTTPClient(auth=None, session=Mock())
    client.values_get = grid.values_get  # type: ignore[assignment]
    client.batch_update = grid.batch_update  # type: ignore[assignment]
    properties = {
        "sheetId": SHEET_ID,
        "title": TITLE,
        "gridProperties": {"rowCount": first_row + len(ids) + 10, "columnCount": 5},
    }
    worksheet = Worksheet(
        spreadsheet=Mock(),
        properties=properties,
        spreadsheet_id=SPREADSHEET_ID,
        client=client,
    )
    return worksheet, grid


def a1_for(ids: Sequence[Any], first_row: int = 2) -> str:
    return "A{}:A{}".format(first_row, first_row + len(ids) - 1)


class PlanValidationTest(unittest.TestCase):
    def test_blank_id_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "", "c"])

        with self.assertRaises(MissingIDsError) as caught:
            plan_row_reorder(worksheet, "A2:A4", ["a", "", "c"])

        self.assertEqual(caught.exception.blank_rows, [3])
        self.assertIsInstance(caught.exception, InvalidInputValue)

    def test_trailing_blank_is_rejected(self):
        # Sheets omits trailing blanks; the planner must pad and reject them.
        worksheet, _ = make_worksheet(["a", "b", ""])

        with self.assertRaises(MissingIDsError) as caught:
            plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])

        self.assertEqual(caught.exception.blank_rows, [4])

    def test_duplicate_in_worksheet_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b", "a"])

        with self.assertRaises(DuplicateIDsError) as caught:
            plan_row_reorder(worksheet, "A2:A4", ["a", "b", "a"])

        self.assertEqual(caught.exception.duplicates, ["a"])
        self.assertEqual(caught.exception.source, "worksheet")

    def test_duplicate_in_target_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b", "c"])

        with self.assertRaises(DuplicateIDsError) as caught:
            plan_row_reorder(worksheet, "A2:A4", ["a", "b", "b"])

        self.assertEqual(caught.exception.duplicates, ["b"])
        self.assertEqual(caught.exception.source, "target_order")

    def test_set_mismatch_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b", "c"])

        with self.assertRaises(IDSetMismatchError) as caught:
            plan_row_reorder(worksheet, "A2:A4", ["a", "b", "d"])

        self.assertEqual(caught.exception.only_in_worksheet, ["c"])
        self.assertEqual(caught.exception.only_in_target, ["d"])

    def test_membership_alone_is_not_enough(self):
        # Same members, different sizes: the target may not be a permutation
        # just because the sets overlap.
        worksheet, _ = make_worksheet(["a", "b", "c", "d"])

        with self.assertRaises((DuplicateIDsError, IDSetMismatchError)):
            plan_row_reorder(worksheet, "A2:A5", ["a", "b", "c"])

    def test_multi_column_region_reads_only_the_id_column(self):
        worksheet, client = make_worksheet(["b", "a"])

        plan = plan_row_reorder(worksheet, "A2:D3", ["a", "b"])

        self.assertEqual(plan.minimum_moves, 1)
        self.assertEqual(client.value_gets[0]["range"], "'Reconcile'!A2:A3")

    def test_range_must_be_bounded(self):
        worksheet, _ = make_worksheet(["a", "b"])

        for range_name in ("A2:A", "A:A", "2:3"):
            with self.assertRaises(ValueError):
                plan_row_reorder(worksheet, range_name, ["a", "b"])

    def test_single_cell_range_is_a_one_row_region(self):
        worksheet, client = make_worksheet(["a", "b"])

        plan = plan_row_reorder(worksheet, "A2", ["a"])

        self.assertTrue(plan.is_noop)
        self.assertEqual(plan.row_count, 1)
        self.assertEqual(client.value_gets[0]["range"], "'Reconcile'!A2:A2")

    def test_range_title_must_match(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "Other!A2:A3", ["a", "b"])

    def test_range_may_name_the_worksheet(self):
        worksheet, client = make_worksheet(["b", "a"])

        plan = plan_row_reorder(worksheet, "Reconcile!A2:A3", ["a", "b"])

        self.assertEqual(plan.minimum_moves, 1)
        self.assertEqual(client.value_gets[0]["range"], "'Reconcile'!A2:A3")

    def test_empty_target_is_rejected(self):
        worksheet, _ = make_worksheet(["a"])

        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "A2:A2", [])

    def test_bare_string_target_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "A2:A3", "ab")

    def test_range_beyond_sheet_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(
                worksheet, "A2:A{}".format(worksheet.row_count + 1), ["a", "b"]
            )

    def test_id_column_outside_region_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "A2:B3", ["a", "b"], id_column="D")
        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "A2:B3", ["a", "b"], id_column=9)

    def test_id_column_range_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b"])

        for bad in ("A:B", "A2:B3", ""):
            with self.assertRaises(ValueError):
                plan_row_reorder(worksheet, "A2:B3", ["a", "b"], id_column=bad)

    def test_id_column_must_be_parsable(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(worksheet, "A2:B3", ["a", "b"], id_column="1A")
        with self.assertRaises(ValueError):
            plan_row_reorder(
                worksheet, "A2:B3", ["a", "b"], id_column=1.5  # type: ignore[arg-type]
            )

    def test_boolean_id_column_is_rejected(self):
        worksheet, _ = make_worksheet(["a", "b"])

        with self.assertRaises(ValueError):
            plan_row_reorder(
                worksheet, "A2:B3", ["a", "b"], id_column=True  # type: ignore[arg-type]
            )


class IdColumnTest(unittest.TestCase):
    def setUp(self):
        # IDs live in column C, with dummy columns on either side.
        self.rows = [
            ["note-1", "x", "b"],
            ["note-2", "y", "a"],
            ["note-3", "z", "c"],
        ]
        self.worksheet, self.grid = make_worksheet(
            [], first_row=2, value_column=self.rows
        )

    def _plan(self, id_column):
        return plan_row_reorder(
            self.worksheet, "A2:C4", ["a", "b", "c"], id_column=id_column
        )

    def test_defaults_to_the_first_column_of_the_region(self):
        worksheet, client = make_worksheet(
            ["b", "a"], value_column=[["b", "z"], ["a", "z"]]
        )

        plan = plan_row_reorder(worksheet, "A2:B3", ["a", "b"])

        self.assertEqual(plan.id_column, 1)
        self.assertEqual(client.value_gets[0]["range"], "'Reconcile'!A2:A3")
        self.assertEqual(plan.minimum_moves, 1)

    def test_column_number_selects_the_id_column(self):
        plan = self._plan(3)

        self.assertEqual(plan.id_column, 3)
        self.assertEqual(plan.current_ids, ("b", "a", "c"))
        self.assertEqual(plan.minimum_moves, 1)
        self.assertEqual(self.grid.value_gets[0]["range"], "'Reconcile'!C2:C4")

    def test_column_letter_and_label_select_the_id_column(self):
        for id_column in ("C", "c", "C1", "C99"):
            with self.subTest(id_column=id_column):
                worksheet, grid = make_worksheet(
                    [], first_row=2, value_column=[list(row) for row in self.rows]
                )
                plan = plan_row_reorder(
                    worksheet, "A2:C4", ["a", "b", "c"], id_column=id_column
                )
                self.assertEqual(plan.id_column, 3)
                self.assertEqual(grid.value_gets[0]["range"], "'Reconcile'!C2:C4")

    def test_numeric_string_selects_the_id_column(self):
        plan = self._plan("3")

        self.assertEqual(plan.id_column, 3)

    def test_other_columns_are_ignored_when_reading_ids(self):
        plan = self._plan(3)
        plan.execute()

        # Whole rows move, so the note columns follow the ID they belong to.
        self.assertEqual(
            self.grid.rows,
            [
                ["note-2", "y", "a"],
                ["note-1", "x", "b"],
                ["note-3", "z", "c"],
            ],
        )


class NoOpTest(unittest.TestCase):
    def test_already_ordered_sheet_is_not_written(self):
        worksheet, client = make_worksheet(["a", "b", "c"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])

        self.assertTrue(plan.is_noop)
        self.assertEqual(plan.moves, ())
        self.assertEqual(len(plan), 0)
        self.assertIsNone(plan.execute())
        self.assertEqual(client.batch_calls, [])
        self.assertIn("zero moves", plan.explain_minimality())
        self.assertIn("already in the requested order", plan.preview())


class PlanTest(unittest.TestCase):
    def test_single_swap_uses_at_most_one_move(self):
        worksheet, client = make_worksheet(["a", "b"])

        plan = plan_row_reorder(worksheet, "A2:A3", ["b", "a"])

        self.assertEqual(plan.minimum_moves, 1)
        plan.execute()
        self.assertEqual(client.id_column(), ["b", "a"])

    def test_reverse_order(self):
        worksheet, client = make_worksheet(["a", "b", "c", "d"])

        plan = plan_row_reorder(worksheet, "A2:A5", ["d", "c", "b", "a"])
        plan.execute()

        self.assertEqual(client.id_column(), ["d", "c", "b", "a"])

    def test_whole_rows_travel_with_their_id(self):
        rows = [
            ["h1", "h2"],
            ["c", "payload-c"],
            ["a", "payload-a"],
            ["b", "payload-b"],
            ["total", "sum"],
        ]
        worksheet, client = make_worksheet(
            ["c", "a", "b"], first_row=1, value_column=rows
        )

        plan = plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])
        plan.execute()

        self.assertEqual(
            client.rows,
            [
                ["h1", "h2"],
                ["a", "payload-a"],
                ["b", "payload-b"],
                ["c", "payload-c"],
                ["total", "sum"],
            ],
        )

    def test_rows_outside_the_region_never_move(self):
        rows = [
            ["header"],
            ["c"],
            ["a"],
            ["b"],
            ["summary"],
        ]
        worksheet, client = make_worksheet(
            ["c", "a", "b"], first_row=1, value_column=rows
        )

        plan = plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])
        plan.execute()

        self.assertEqual(client.rows[0], ["header"])
        self.assertEqual(client.rows[-1], ["summary"])
        # Every request is confined to the data region [1, 4).
        for request in client.batch_calls[0]["requests"]:
            source = request["moveDimension"]["source"]
            self.assertGreaterEqual(source["startIndex"], 1)
            self.assertLess(source["endIndex"], 4)
            destination = request["moveDimension"]["destinationIndex"]
            self.assertGreaterEqual(destination, 1)
            self.assertLessEqual(destination, 4)

    def test_moves_are_sent_in_one_batch(self):
        worksheet, client = make_worksheet(["e", "d", "c", "b", "a"])

        plan = plan_row_reorder(worksheet, "A2:A6", ["a", "b", "c", "d", "e"])
        plan.execute()

        self.assertEqual(len(client.batch_calls), 1)
        self.assertEqual(len(client.batch_calls[0]["requests"]), len(plan))
        plan.execute()  # executing again replays the same plan on new state
        self.assertEqual(len(client.batch_calls), 2)

    def test_plan_is_immutable_preview_does_not_write(self):
        worksheet, client = make_worksheet(["c", "b", "a"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])

        with self.assertRaises(AttributeError):
            plan.range_name = "B2:B4"  # type: ignore[misc]

        self.assertIn("Reorder plan", plan.preview())
        self.assertEqual(client.batch_calls, [])

    def test_execute_result_is_returned(self):
        worksheet, client = make_worksheet(["b", "a"])

        plan = plan_row_reorder(worksheet, "A2:A3", ["a", "b"])
        result = plan.execute()

        self.assertEqual(result, {"replies": [{}]})

    def test_reorder_rows_wrapper_plans_and_executes(self):
        worksheet, client = make_worksheet(["c", "a", "b"])

        result = worksheet.reorder_rows("A2:A4", ["a", "b", "c"])

        self.assertEqual(result, {"replies": [{}]})
        self.assertEqual(client.id_column(), ["a", "b", "c"])

    def test_reorder_rows_wrapper_is_a_noop_when_already_ordered(self):
        worksheet, client = make_worksheet(["a", "b"])

        self.assertIsNone(worksheet.reorder_rows("A2:A3", ["a", "b"]))
        self.assertEqual(client.batch_calls, [])

    def test_module_level_reorder_rows(self):
        from gspread.reorder import reorder_rows

        worksheet, client = make_worksheet(["b", "a"])

        reorder_rows(worksheet, "A2:A3", ["a", "b"])

        self.assertEqual(client.id_column(), ["a", "b"])

    def test_preview_lists_each_move_with_coordinates(self):
        worksheet, _ = make_worksheet(["c", "a", "b"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["a", "b", "c"])
        text = plan.preview()

        self.assertIn("Reorder plan for worksheet 'Reconcile'", text)
        self.assertIn("ID range A2:A4", text)
        self.assertIn("Minimum single-row moves: 1", text)
        self.assertIn("ID 'c': row 2 -> row 4 (down)", text)
        self.assertIn("moveDimension source rows [1, 2) destinationIndex 4", text)

    def test_kept_ids_describe_unmoved_rows(self):
        worksheet, _ = make_worksheet(["a", "b", "c", "d", "e"])

        plan = plan_row_reorder(worksheet, "A2:A6", ["a", "c", "e", "b", "d"])

        # Target ranks in current order: 0, 3, 1, 4, 2. Any of several
        # length-3 increasing runs is valid; the kept IDs must be one of them
        # and must never appear among the moved rows.
        self.assertEqual(len(plan.kept_ids), lis_length([0, 3, 1, 4, 2]))
        self.assertFalse(set(plan.kept_ids) & {move.id for move in plan.moves})
        # A valid kept set keeps its relative order in both layouts.
        current_positions = [plan.current_ids.index(id_) for id_ in plan.kept_ids]
        target_positions = [plan.target_ids.index(id_) for id_ in plan.kept_ids]
        self.assertEqual(current_positions, sorted(current_positions))
        self.assertEqual(target_positions, sorted(target_positions))


class MoveSemanticsTest(unittest.TestCase):
    def test_requests_use_pre_removal_destination(self):
        # Moving "c" up above "a": it ends at index 1, and destinationIndex 1
        # already means "before the current row 2" in pre-removal
        # coordinates, i.e. no off-by-one adjustment for upward moves.
        worksheet, _ = make_worksheet(["a", "b", "c"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["c", "a", "b"])

        self.assertEqual(len(plan.moves), 1)
        move = plan.moves[0]
        self.assertEqual(move.source_row, 4)
        self.assertEqual(move.final_row, 2)
        self.assertEqual(move.destination_row, 2)
        self.assertEqual(move.direction, "up")
        self.assertEqual(
            move.to_request(SHEET_ID),
            {
                "moveDimension": {
                    "source": {
                        "sheetId": SHEET_ID,
                        "dimension": "ROWS",
                        "startIndex": 3,
                        "endIndex": 4,
                    },
                    "destinationIndex": 1,
                }
            },
        )

    def test_moving_down_adds_one_to_the_final_index(self):
        # Moving "a" below "c" lands at index 3, but Sheets reads
        # destinationIndex in pre-removal coordinates, so it must be 4.
        worksheet, _ = make_worksheet(["a", "b", "c", "d"])

        plan = plan_row_reorder(worksheet, "A2:A5", ["b", "c", "a", "d"])

        move = plan.moves[0]
        self.assertEqual(move.id, "a")
        self.assertEqual(move.source_row, 2)
        self.assertEqual(move.final_row, 4)
        self.assertEqual(move.destination_row, 5)
        self.assertEqual(move.direction, "down")
        self.assertEqual(
            move.to_request(SHEET_ID)["moveDimension"]["destinationIndex"], 4
        )

    def test_moving_up_uses_destination_directly(self):
        worksheet, _ = make_worksheet(["a", "b", "c"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["c", "a", "b"])

        move = plan.moves[0]
        self.assertEqual(move.id, "c")
        self.assertEqual(move.source_row, 4)
        self.assertEqual(move.final_row, 2)
        self.assertEqual(move.destination_row, 2)
        self.assertEqual(move.direction, "up")

    def test_plan_matches_sheet_semantics_for_every_small_permutation(self):
        for n in range(1, 7):
            base = ["id-{}".format(index) for index in range(n)]
            for permutation in itertools.permutations(base):
                with self.subTest(target=permutation):
                    worksheet, client = make_worksheet(base)
                    plan = plan_row_reorder(worksheet, a1_for(base), list(permutation))
                    plan.execute()
                    self.assertEqual(client.id_column(), list(permutation))

    def test_target_row_is_the_resting_place_after_the_whole_plan(self):
        # For every small permutation, each moved ID must be reported at the
        # row it really occupies once all moves have run, and that must be its
        # row in the target order.
        for n in range(1, 7):
            base = ["id-{}".format(index) for index in range(n)]
            for permutation in itertools.permutations(base):
                with self.subTest(target=permutation):
                    worksheet, _ = make_worksheet(base)
                    plan = plan_row_reorder(worksheet, a1_for(base), list(permutation))

                    resting = {id_: row for row, id_ in enumerate(permutation, 2)}
                    for move in plan.moves:
                        self.assertEqual(move.target_row, resting[move.id])
                        self.assertEqual(
                            move.target_row, plan.first_row + move.target_index
                        )

    def test_target_row_can_differ_from_the_row_after_the_move(self):
        # Moving "3" up places it on row 4, then moving "0" to the bottom
        # shifts "3" one row further up. Only target_row names its resting
        # place; final_row is the position right after that one request.
        worksheet, _ = make_worksheet(["0", "1", "2", "3"])

        plan = plan_row_reorder(worksheet, "A2:A5", ["1", "3", "2", "0"])

        move = plan.moves[0]
        self.assertEqual(move.id, "3")
        self.assertEqual(move.final_row, 4)
        self.assertEqual(move.target_row, 3)
        self.assertEqual(move.net_direction, "up")
        self.assertIn("ID '3': row 5 -> row 3 (up)", plan.preview())

    def test_a_row_can_move_and_end_up_on_its_own_row(self):
        # "2" has to leave its row so the descending run can pass it, then
        # ends up back on row 4: it is still a moved row, and the plan must
        # not pretend it stayed put.
        worksheet, client = make_worksheet(["0", "1", "2", "3", "4"])

        plan = plan_row_reorder(worksheet, "A2:A6", ["3", "4", "2", "1", "0"])

        move = next(move for move in plan.moves if move.id == "2")
        self.assertEqual((move.source_row, move.target_row), (4, 4))
        self.assertEqual(move.net_direction, "in place")
        self.assertIn("ID '2': row 4 -> row 4 (in place)", plan.preview())
        plan.execute()
        self.assertEqual(client.id_column(), ["3", "4", "2", "1", "0"])


class MinimalityTest(unittest.TestCase):
    def test_move_count_equals_n_minus_lis(self):
        for n in range(1, 7):
            base = list(range(n))
            for permutation in itertools.permutations(base):
                with self.subTest(target=permutation):
                    worksheet, client = make_worksheet(base)
                    plan = plan_row_reorder(worksheet, a1_for(base), list(permutation))
                    expected = n - lis_length(list(permutation))
                    self.assertEqual(plan.minimum_moves, expected)

                    plan.execute()
                    self.assertEqual(client.id_column(), list(permutation))

    def test_five_rows_first_last_swap(self):
        # Swapping the ends of five rows leaves a 3-row increasing block, so
        # two moves are enough even though both moved rows travel far.
        worksheet, client = make_worksheet(["a", "b", "c", "d", "e"])

        plan = plan_row_reorder(worksheet, "A2:A6", ["e", "b", "c", "d", "a"])

        self.assertEqual(plan.minimum_moves, 2)
        self.assertEqual(len(plan), 2)
        plan.execute()
        self.assertEqual(client.id_column(), ["e", "b", "c", "d", "a"])

    def test_each_row_moves_at_most_once(self):
        ids = ["r{}".format(index) for index in range(30)]
        target = ids[::-1]

        worksheet, client = make_worksheet(ids)
        plan = plan_row_reorder(worksheet, a1_for(ids), target)

        self.assertEqual(len({move.id for move in plan.moves}), len(plan.moves))
        self.assertEqual(plan.minimum_moves, len(ids) - 1)
        plan.execute()
        self.assertEqual(client.id_column(), target)

    def test_explanation_counts_arguments(self):
        worksheet, _ = make_worksheet(["a", "b", "c", "d"])

        plan = plan_row_reorder(worksheet, "A2:A5", ["b", "a", "d", "c"])
        text = plan.explain_minimality()

        self.assertIn("common subsequence", text)
        self.assertIn("contains 2 of 4 rows", text)
        self.assertIn("4 - 2 = 2 moves", text)
        self.assertIn("moves exactly 2 row(s)", text)

    def test_explanation_truncates_long_kept_lists(self):
        ids = ["id{}".format(index) for index in range(20)]

        worksheet, _ = make_worksheet(ids)
        # Rotating the first five rows to the end keeps a 15-row run intact,
        # which is long enough to truncate the preview list.
        target = ids[5:] + ids[:5]
        plan = plan_row_reorder(worksheet, a1_for(ids), target)
        text = plan.explain_minimality()

        self.assertIn("…", text)
        self.assertIn("15 of 20 rows", text)
        self.assertIn("5 moves", text)

    def test_kept_rows_prefer_already_correct_positions(self):
        # Target ranks are 0, 1, 4, 3, 2. Two length-3 runs exist: (a, b, d)
        # with "d" already in place, or (a, b, e) where "e" must move.
        # The planner keeps the rows that need no move.
        worksheet, client = make_worksheet(["a", "b", "c", "d", "e"])

        plan = plan_row_reorder(worksheet, "A2:A6", ["a", "b", "e", "d", "c"])

        self.assertEqual(plan.kept_ids, ("a", "b", "d"))
        self.assertEqual(plan.minimum_moves, 2)
        plan.execute()
        self.assertEqual(client.id_column(), ["a", "b", "e", "d", "c"])


class IdentityNormalisationTest(unittest.TestCase):
    def test_integers_match_unformatted_numeric_cells(self):
        worksheet, client = make_worksheet([101, 102, 103])

        plan = plan_row_reorder(worksheet, "A2:A4", [103, 101, 102])
        plan.execute()

        self.assertEqual(client.id_column(), [103, 101, 102])

    def test_integral_floats_compare_equal_to_ints(self):
        worksheet, _ = make_worksheet([1, 2, 3])

        plan = plan_row_reorder(worksheet, "A2:A4", [3.0, 2.0, 1.0])

        self.assertEqual(plan.target_ids, ("3", "2", "1"))
        self.assertEqual(plan.minimum_moves, 2)

    def test_unformatted_is_the_default_render_option(self):
        worksheet, client = make_worksheet(["a", "b"])

        plan_row_reorder(worksheet, "A2:A3", ["b", "a"])

        self.assertEqual(
            client.value_gets[0]["params"]["valueRenderOption"],
            ValueRenderOption.unformatted,
        )
        self.assertEqual(client.value_gets[0]["params"]["majorDimension"], "COLUMNS")

    def test_property_accessors(self):
        worksheet, _ = make_worksheet(["a", "b", "c"])

        plan = plan_row_reorder(worksheet, "A2:A4", ["b", "a", "c"])

        self.assertIsInstance(plan, RowReorderPlan)
        self.assertEqual(plan.sheet_title, TITLE)
        self.assertEqual(plan.row_count, 3)
        self.assertEqual((plan.first_row, plan.last_row, plan.column), (2, 4, 1))
        self.assertEqual(plan.current_ids, ("a", "b", "c"))
        self.assertEqual(plan.target_ids, ("b", "a", "c"))
        self.assertEqual(list(plan), list(plan.moves))
