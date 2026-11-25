"""
gspread.reorder
~~~~~~~~~~~~~~~

Plan and apply row reorders in a worksheet using a column of unique IDs.

The planner reads an ID column from a contiguous data region, checks that the
requested target order is exactly a permutation of the current IDs, and
computes the smallest possible number of single-row ``moveDimension``
requests. Rows outside the ID region (headers above it, summary rows below
it) are never used as a source or a destination inside the region, and because
whole rows are moved, formulas, notes and formatting travel with their row.

.. versionadded:: 6.3.0
"""

from collections import Counter
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Dict,
    Iterable,
    List,
    Optional,
    Sequence,
    Tuple,
    Union,
)

from .exceptions import (
    DuplicateIDsError,
    IDSetMismatchError,
    IncorrectCellLabel,
    InvalidInputValue,
    MissingIDsError,
)
from .utils import (
    CELL_ADDR_RE,
    Dimension,
    ValueRenderOption,
    _a1_to_rowcol_unbounded,
    a1_range_to_grid_range,
    a1_to_rowcol,
    absolute_range_name,
    extract_title_from_range,
    get_a1_from_absolute_range,
    rowcol_to_a1,
)

if TYPE_CHECKING is True:
    from .worksheet import Worksheet


#: Number of kept IDs shown in :meth:`RowReorderPlan.explain_minimality`.
_MAX_IDS_IN_EXPLANATION = 8


def _id_text(value: Any) -> str:
    """Return the canonical text form used to compare IDs.

    Values read from a sheet are usually strings already. Numbers supplied by
    the caller (for example ``1`` or ``1.0``) are normalised so that they can
    be compared with numeric cells read with
    :attr:`~gspread.utils.ValueRenderOption.unformatted`.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _is_blank(value: Any) -> bool:
    return value is None or value == ""


def _strip_sheet_title(worksheet: "Worksheet", range_name: str) -> str:
    """Return the A1 part of ``range_name``, checking any title it names."""
    if "!" not in range_name:
        return range_name

    try:
        title = extract_title_from_range(range_name)
    except InvalidInputValue as ex:
        raise ValueError(
            "cannot parse worksheet title from range_name {!r}".format(range_name)
        ) from ex
    if title != worksheet.title:
        raise ValueError(
            "range_name {!r} refers to worksheet {!r}, not {!r}".format(
                range_name, title, worksheet.title
            )
        )
    return get_a1_from_absolute_range(range_name)


def _parse_data_range(
    worksheet: "Worksheet", range_name: str
) -> Tuple[str, int, int, int, int]:
    """Validate ``range_name`` and return the region's bounds.

    Returns ``(a1_range, first_row, last_row, first_column, last_column)``
    with 1-based, inclusive row and column numbers. ``range_name`` must be
    bounded on all sides and must lie inside ``worksheet``.
    """
    if not isinstance(range_name, str) or not range_name.strip():
        raise ValueError("range_name must be a non-empty A1 notation string")

    a1_range = _strip_sheet_title(worksheet, range_name)

    try:
        grid_range = a1_range_to_grid_range(a1_range)
    except InvalidInputValue as ex:
        raise ValueError("cannot parse range_name {!r}".format(range_name)) from ex

    bounds = (
        "startRowIndex",
        "endRowIndex",
        "startColumnIndex",
        "endColumnIndex",
    )
    if any(bound not in grid_range for bound in bounds):
        raise ValueError(
            "range_name {!r} must be bounded on all sides, "
            "for example 'A2:A100'".format(range_name)
        )

    start_row = grid_range["startRowIndex"]
    end_row = grid_range["endRowIndex"]
    if end_row <= start_row:
        raise ValueError(
            "range_name {!r} must contain at least one row".format(range_name)
        )
    if end_row > worksheet.row_count:
        raise ValueError(
            "range_name {!r} exceeds the worksheet size of {} rows".format(
                range_name, worksheet.row_count
            )
        )

    start_column = grid_range["startColumnIndex"]
    end_column = grid_range["endColumnIndex"]
    if end_column > worksheet.col_count:
        raise ValueError(
            "range_name {!r} exceeds the worksheet size of {} columns".format(
                range_name, worksheet.col_count
            )
        )

    return a1_range, start_row + 1, end_row, start_column + 1, end_column


def _column_number_from_label(label: str) -> int:
    """Convert a column reference such as ``"3"``, ``"C"`` or ``"C1"`` to a number."""
    if label.isdigit():
        return int(label)
    if CELL_ADDR_RE.match(label) and label[-1].isdigit():
        column: Any = a1_to_rowcol(label)[1]
    else:
        column = _a1_to_rowcol_unbounded(label)[1]
    if not isinstance(column, int):
        raise IncorrectCellLabel(label)
    return column


def _resolve_id_column(
    range_name: str,
    first_column: int,
    last_column: int,
    id_column: Union[int, str, None],
) -> int:
    """Return the 1-based worksheet column that holds the IDs.

    ``id_column`` may be a 1-based column number inside the region, a single
    column letter, or an A1 label such as ``"C1"``. ``None`` (the default)
    means the region's first column.
    """
    if id_column is None:
        return first_column

    if isinstance(id_column, str):
        label = id_column.strip()
        if not label:
            raise ValueError("id_column must not be an empty string")
        if ":" in label or "!" in label:
            raise ValueError(
                "id_column {!r} must name a single column, not a range".format(
                    id_column
                )
            )
        try:
            column = _column_number_from_label(label)
        except (IncorrectCellLabel, ValueError) as ex:
            raise ValueError(
                "cannot parse id_column {!r} as a column".format(id_column)
            ) from ex
    elif isinstance(id_column, int) and not isinstance(id_column, bool):
        column = id_column
    else:
        raise ValueError(
            "id_column must be a column number, a column letter or None, "
            "not {!r}".format(type(id_column).__name__)
        )

    if not first_column <= column <= last_column:
        raise ValueError(
            "id_column {} is outside range_name {!r} (columns {} to {})".format(
                column, range_name, first_column, last_column
            )
        )
    return column


def _read_id_column(
    worksheet: "Worksheet",
    first_row: int,
    last_row: int,
    column: int,
    value_render_option: ValueRenderOption,
) -> List[Any]:
    """Read only the ID column of the region, padding blank trailing cells.

    ``majorDimension=COLUMNS`` returns one list per requested column, in which
    interior blank cells keep their position, so blanks can be reported at the
    right row. Missing trailing cells are padded so the returned list always
    has exactly ``last_row - first_row + 1`` entries.
    """
    rows = last_row - first_row + 1
    data = worksheet.client.values_get(
        worksheet.spreadsheet_id,
        absolute_range_name(
            worksheet.title,
            "{}:{}".format(
                rowcol_to_a1(first_row, column), rowcol_to_a1(last_row, column)
            ),
        ),
        params={
            "majorDimension": Dimension.cols,
            "valueRenderOption": value_render_option,
        },
    )
    values = data.get("values") or []
    column_values: List[Any] = list(values[0]) if values else []
    if len(column_values) > rows:
        column_values = column_values[:rows]
    column_values.extend([""] * (rows - len(column_values)))
    return column_values


class _FenwickMax:
    """Fenwick tree over 1-based positions holding the best ``(length, in_place)``.

    Each entry also stores the data position that produced it, so the
    subsequence can be walked back once the best chain is known.
    """

    def __init__(self, size: int):
        self.size = size
        self.tree: List[Tuple[int, int, int]] = [(0, 0, -1)] * (size + 1)

    @staticmethod
    def _better(candidate: Tuple[int, int, int], current: Tuple[int, int, int]) -> bool:
        return candidate[:2] > current[:2]

    def best_before(self, position: int) -> Tuple[int, int, int]:
        """Best entry for positions strictly smaller than ``position``."""
        best: Tuple[int, int, int] = (0, 0, -1)
        while position > 0:
            if self._better(self.tree[position], best):
                best = self.tree[position]
            position -= position & -position
        return best

    def add(self, position: int, value: Tuple[int, int, int]) -> None:
        while position <= self.size:
            if self._better(value, self.tree[position]):
                self.tree[position] = value
            position += position & -position


def _kept_row_positions(ranks: Sequence[int]) -> List[int]:
    """Return positions of a maximum increasing subsequence of ``ranks``.

    ``ranks[i]`` is the target position of the ID currently at data position
    ``i``. Unmoved rows keep their relative order, so they must correspond to
    an increasing subsequence of ``ranks``. Among the longest subsequences the
    one with the most rows already at their final position is preferred; this
    keeps already-correct rows out of the plan whenever that costs no extra
    moves.

    The implementation is a Fenwick tree over target positions storing
    ``(length, already_in_place)`` pairs, so it runs in ``O(n log n)``.
    """
    n = len(ranks)
    if n == 0:
        return []

    tree = _FenwickMax(n)
    previous = [-1] * n
    lengths = [0] * n
    in_place = [0] * n

    for index, rank in enumerate(ranks):
        best_length, best_in_place, best_index = tree.best_before(rank)
        lengths[index] = best_length + 1
        in_place[index] = best_in_place + (1 if rank == index else 0)
        previous[index] = best_index
        tree.add(rank + 1, (lengths[index], in_place[index], index))

    end = max(range(n), key=lambda index: (lengths[index], in_place[index]))

    picked: List[int] = []
    while end != -1:
        picked.append(end)
        end = previous[end]
    picked.reverse()
    return picked


def _build_moves(
    current_ids: Sequence[str],
    target_ids: Sequence[str],
    kept_positions: Sequence[int],
    first_row: int,
) -> List["RowMove"]:
    """Build single-row moves using Sheets' pre-removal coordinates.

    Each move's ``sourceIndex`` is the row's position *at the moment the
    request runs*, and ``destinationIndex`` follows the Sheets semantics: it is
    interpreted against the coordinates before the source row is removed. The
    plan is computed by simulating the sheet after every move.
    """
    index_of = {id_: index for index, id_ in enumerate(current_ids)}
    kept = set(kept_positions)
    order = list(range(len(current_ids)))
    moves: List["RowMove"] = []

    for target_position, id_ in enumerate(target_ids):
        origin = index_of[id_]
        if origin in kept:
            continue

        source = order.index(origin)

        if target_position == 0:
            final = 0
        else:
            previous_origin = index_of[target_ids[target_position - 1]]
            previous_position = order.index(previous_origin)
            if previous_position < source:
                final = previous_position + 1
            else:
                final = previous_position

        if final == source:
            # Mathematically unreachable for a maximum kept subsequence: it
            # would mean one more row can stay in place, contradicting
            # maximality. Guarding keeps the plan valid if that ever changes.
            kept.add(origin)
            continue

        # destinationIndex is interpreted in pre-removal coordinates:
        # moving down needs the index *after* the final position, moving up
        # uses the final position directly.
        destination = final if final < source else final + 1

        order.pop(source)
        order.insert(final, origin)

        moves.append(
            RowMove(
                id=id_,
                target_index=target_position,
                source_index=first_row - 1 + source,
                destination_index=first_row - 1 + destination,
                final_index=first_row - 1 + final,
                target_row=first_row + target_position,
            )
        )

    resulting_ids = [current_ids[origin] for origin in order]
    if resulting_ids != list(target_ids):  # pragma: no cover - invariant
        raise RuntimeError(
            "internal error: reorder plan does not produce the target order"
        )

    return moves


@dataclass(frozen=True)
class RowMove:
    """A single ``moveDimension`` request moving one row.

    All indexes are zero-based, matching the Sheets API. ``source_index`` is
    the row's position when this request runs; ``destination_index`` is
    interpreted by Sheets against the coordinates *before* the source row is
    removed; ``final_index`` is where the row ends up right after this request.
    Later moves in the same plan can still shift it, so the row's final resting
    place is ``target_row`` (1-based, also available as ``target_index``
    zero-based).

    ``target_index`` is the row's position in the requested order.
    """

    id: str
    target_index: int
    source_index: int
    destination_index: int
    final_index: int
    target_row: int

    @property
    def source_row(self) -> int:
        """Source row number, 1-based (as shown in the Sheets UI)."""
        return self.source_index + 1

    @property
    def destination_row(self) -> int:
        """Destination row number, 1-based, in pre-removal coordinates."""
        return self.destination_index + 1

    @property
    def final_row(self) -> int:
        """Row number right after this move, 1-based.

        Later moves in the plan may still shift the row; use
        :attr:`target_row` for its resting place once the whole plan has run.
        """
        return self.final_index + 1

    @property
    def direction(self) -> str:
        """``"down"`` if this request lands the row lower, else ``"up"``.

        This is the immediate effect of the request. Later moves in the plan
        can push the row further; :attr:`net_direction` describes the overall
        change from ``source_row`` to :attr:`target_row`.
        """
        return "down" if self.final_index > self.source_index else "up"

    @property
    def net_direction(self) -> str:
        """``"down"``/``"up"``/``"in place"`` from ``source_row`` to ``target_row``.

        A row can move away and come back to the row it started on while other
        rows pass it; ``"in place"`` reports that it ends up where it began.
        """
        if self.target_row > self.source_row:
            return "down"
        if self.target_row < self.source_row:
            return "up"
        return "in place"

    def to_request(self, sheet_id: int) -> Dict[str, Any]:
        """Return the Sheets API ``moveDimension`` request for this move."""
        return {
            "moveDimension": {
                "source": {
                    "sheetId": sheet_id,
                    "dimension": Dimension.rows.value,
                    "startIndex": self.source_index,
                    "endIndex": self.source_index + 1,
                },
                "destinationIndex": self.destination_index,
            }
        }


@dataclass(frozen=True)
class RowReorderPlan:
    """A previewable, executable plan that reorders rows by ID.

    Get an instance from :meth:`gspread.worksheet.Worksheet.plan_row_reorder`.
    The plan is computed once from the current sheet contents. Use
    :meth:`preview` to inspect it, :meth:`to_requests` to inspect the raw
    Sheets requests, and :meth:`execute` to send them all in one
    ``batchUpdate`` call. An already-correct order produces an empty plan and
    :meth:`execute` does not call the API at all.
    """

    worksheet: "Worksheet" = field(repr=False, compare=False)
    range_name: str
    a1_range: str
    id_column: int
    first_row: int
    last_row: int
    current_ids: Tuple[str, ...]
    target_ids: Tuple[str, ...]
    kept_ids: Tuple[str, ...]
    moves: Tuple[RowMove, ...]

    @property
    def sheet_title(self) -> str:
        """Title of the worksheet this plan applies to."""
        return self.worksheet.title

    @property
    def row_count(self) -> int:
        """Number of rows covered by the data range."""
        return self.last_row - self.first_row + 1

    @property
    def column(self) -> int:
        """1-based worksheet column that holds the IDs."""
        return self.id_column

    @property
    def first_row_index(self) -> int:
        """Zero-based index of the first row of the data region."""
        return self.first_row - 1

    @property
    def last_row_index(self) -> int:
        """Zero-based index of the last row of the data region."""
        return self.last_row - 1

    @property
    def is_noop(self) -> bool:
        """True when the IDs already appear in the requested order."""
        return not self.moves

    @property
    def minimum_moves(self) -> int:
        """Number of single-row moves in the plan (the minimum possible)."""
        return len(self.moves)

    def __len__(self) -> int:
        return len(self.moves)

    def __iter__(self) -> Iterable[RowMove]:
        return iter(self.moves)

    def to_requests(self) -> List[Dict[str, Any]]:
        """Return the ``moveDimension`` requests, in execution order."""
        sheet_id = self.worksheet.id
        return [move.to_request(sheet_id) for move in self.moves]

    def execute(self) -> Optional[Any]:
        """Submit every move in one ``batchUpdate`` call.

        Returns the API response, or ``None`` when the plan is a no-op (in
        which case no request is made).
        """
        if self.is_noop:
            return None
        return self.worksheet.client.batch_update(
            self.worksheet.spreadsheet_id, {"requests": self.to_requests()}
        )

    def explain_minimality(self) -> str:
        """Explain why no plan with fewer single-row moves exists."""
        total = len(self.current_ids)
        moved = len(self.moves)

        if moved == 0:
            return (
                "The IDs already appear in the requested order, so zero moves "
                "are required and no request is sent."
            )

        kept = len(self.kept_ids)
        shown = list(self.kept_ids[:_MAX_IDS_IN_EXPLANATION])
        ids = ", ".join(repr(id_) for id_ in shown)
        if kept > len(shown):
            ids += ", …"

        return (
            "Rows that are never moved keep their relative order, so any set of "
            "unmoved rows must appear in the same order in both the current and "
            "the target layout: it has to be a common subsequence of the two "
            "orders. The longest such subsequence here contains {} of {} rows "
            "({}); every other row must therefore be moved at least once. Since "
            "one moveDimension request relocates exactly one row, no batch of "
            "single-row moves can use fewer than {} - {} = {} moves. This plan "
            "moves exactly {} row(s), each once, so it is optimal.".format(
                kept, total, ids, total, kept, total - kept, moved
            )
        )

    def preview(self) -> str:
        """Return a human-readable preview of the plan and its minimality."""
        lines = [
            "Reorder plan for worksheet {!r}, ID range {} ({} row(s), column {})".format(
                self.sheet_title, self.a1_range, self.row_count, self.column
            )
        ]

        if self.is_noop:
            lines.append(
                "The IDs are already in the requested order: no moves required."
            )
            return "\n".join(lines)

        lines.append(
            "Minimum single-row moves: {} ({} row(s) stay in place)".format(
                len(self.moves), len(self.kept_ids)
            )
        )
        for number, move in enumerate(self.moves, 1):
            lines.append(
                "  {}. ID {!r}: row {} -> row {} ({})".format(
                    number,
                    move.id,
                    move.source_row,
                    move.target_row,
                    move.net_direction,
                )
            )
            lines.append(
                "     moveDimension source rows [{}, {}) destinationIndex {}".format(
                    move.source_index, move.source_index + 1, move.destination_index
                )
            )
        lines.append("")
        lines.append(self.explain_minimality())
        return "\n".join(lines)


def plan_row_reorder(
    worksheet: "Worksheet",
    range_name: str,
    target_order: Sequence[Any],
    id_column: Union[int, str, None] = None,
    value_render_option: ValueRenderOption = ValueRenderOption.unformatted,
) -> RowReorderPlan:
    """Read an ID column, validate it and build a minimum-move reorder plan.

    :param worksheet: The worksheet holding the IDs.
    :param str range_name: A bounded, contiguous range such as ``"A2:D100"``
        (an optional worksheet title is allowed). Its rows are the rows that
        may be moved and its columns are the data area, so the region can
        cover the whole record rather than only the ID column.
    :param target_order: The desired order of the IDs, for example a list of
        strings or numbers. It must contain every ID from the ID column
        exactly once.
    :param id_column: (optional) Which column of ``range_name`` holds the IDs:
        a 1-based column number, a column letter such as ``"A"``, or an A1
        label such as ``"A1"``. Defaults to the first column of the region.
    :param value_render_option: (optional) How sheet values are rendered
        before comparison. Defaults to
        :attr:`~gspread.utils.ValueRenderOption.unformatted`, which compares
        exact stored values; use ``formatted`` to compare displayed text.

    :raises ValueError: If ``range_name`` is not a bounded region of this
        worksheet, ``id_column`` is outside it, or ``target_order`` is empty
        or a plain string.
    :raises MissingIDsError: If a cell in the ID column is blank.
    :raises DuplicateIDsError: If an ID occurs more than once in the sheet or
        in ``target_order``.
    :raises IDSetMismatchError: If the two collections of IDs differ.

    .. versionadded:: 6.3.0
    """
    if isinstance(target_order, str) or not isinstance(target_order, Sequence):
        raise ValueError(
            "target_order must be a sequence of IDs, not {!r}".format(
                type(target_order).__name__
            )
        )
    if len(target_order) == 0:
        raise ValueError("target_order must not be empty")

    a1_range, first_row, last_row, first_column, last_column = _parse_data_range(
        worksheet, range_name
    )
    id_column_number = _resolve_id_column(
        range_name, first_column, last_column, id_column
    )
    raw_ids = _read_id_column(
        worksheet,
        first_row,
        last_row,
        id_column_number,
        value_render_option,
    )

    blank_rows = [
        first_row + offset for offset, value in enumerate(raw_ids) if _is_blank(value)
    ]
    if blank_rows:
        raise MissingIDsError(blank_rows)

    current_ids = [_id_text(value) for value in raw_ids]
    target_ids = [_id_text(value) for value in target_order]

    current_duplicates = sorted(
        id_ for id_, count in Counter(current_ids).items() if count > 1
    )
    if current_duplicates:
        raise DuplicateIDsError(current_duplicates, source="worksheet")

    target_duplicates = sorted(
        id_ for id_, count in Counter(target_ids).items() if count > 1
    )
    if target_duplicates:
        raise DuplicateIDsError(target_duplicates, source="target_order")

    current_set = set(current_ids)
    target_set = set(target_ids)
    only_in_sheet = [id_ for id_ in current_ids if id_ not in target_set]
    only_in_target = [id_ for id_ in target_ids if id_ not in current_set]
    if only_in_sheet or only_in_target:
        raise IDSetMismatchError(only_in_sheet, only_in_target)

    target_positions = {id_: position for position, id_ in enumerate(target_ids)}
    ranks = [target_positions[id_] for id_ in current_ids]
    kept_positions = _kept_row_positions(ranks)
    kept_ids = tuple(current_ids[position] for position in kept_positions)

    moves = _build_moves(current_ids, target_ids, kept_positions, first_row)

    # ``_build_moves`` may (defensively) keep a row it discovers is already in
    # place; recompute the kept set from the actual sources so the plan
    # describes exactly what will happen.
    moved_ids = {move.id for move in moves}
    if len(moved_ids) != len(moves):  # pragma: no cover - invariant
        raise RuntimeError("internal error: an ID would be moved twice")
    kept_ids = tuple(id_ for id_ in current_ids if id_ not in moved_ids)

    return RowReorderPlan(
        worksheet=worksheet,
        range_name=range_name,
        a1_range=a1_range,
        id_column=id_column_number,
        first_row=first_row,
        last_row=last_row,
        current_ids=tuple(current_ids),
        target_ids=tuple(target_ids),
        kept_ids=kept_ids,
        moves=tuple(moves),
    )


def reorder_rows(
    worksheet: "Worksheet",
    range_name: str,
    target_order: Sequence[Any],
    id_column: Union[int, str, None] = None,
    value_render_option: ValueRenderOption = ValueRenderOption.unformatted,
) -> Optional[Any]:
    """Plan and immediately apply a row reorder.

    Convenience wrapper around
    :func:`plan_row_reorder` and
    :meth:`~gspread.reorder.RowReorderPlan.execute`, for callers that do not
    need to inspect the plan first. Prefer the two-step form when a human
    should review the moves. A sheet already in the requested order is left
    untouched and no request is made (``None`` is returned).

    :param worksheet: The worksheet holding the IDs.
    :param str range_name: The bounded data region, such as ``"A2:D100"``.
    :param target_order: The desired order of the IDs.
    :param id_column: (optional) The column of ``range_name`` holding the IDs.
    :param value_render_option: (optional) How cell values are rendered before
        comparison.
    :returns: The ``batchUpdate`` response, or ``None`` for a no-op.

    Example::

        reorder_rows(worksheet, "A2:D100", ordered_ids)

    .. versionadded:: 6.3.0
    """
    return plan_row_reorder(
        worksheet,
        range_name=range_name,
        target_order=target_order,
        id_column=id_column,
        value_render_option=value_render_option,
    ).execute()
