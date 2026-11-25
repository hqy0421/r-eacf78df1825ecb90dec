"""
gspread.exceptions
~~~~~~~~~~~~~~~~~~

Exceptions used in gspread.

"""

from typing import Any, Mapping

from requests import Response


class UnSupportedExportFormat(Exception):
    """Raised when export format is not supported."""


class GSpreadException(Exception):
    """A base class for gspread's exceptions."""


class WorksheetNotFound(GSpreadException):
    """Trying to open non-existent or inaccessible worksheet."""


class NoValidUrlKeyFound(GSpreadException):
    """No valid key found in URL."""


class IncorrectCellLabel(GSpreadException):
    """The cell label is incorrect."""


class InvalidInputValue(GSpreadException):
    """The provided values is incorrect."""


class APIError(GSpreadException):
    """Errors coming from the API itself,
    such as when we attempt to retrieve things that don't exist."""

    def __init__(self, response: Response):
        try:
            error = response.json()["error"]
        except Exception as e:
            # in case we failed to parse the error from the API
            # build an empty error object to notify the caller
            # and keep the exception raise flow running

            error = {
                "code": -1,
                "message": response.text,
                "status": "invalid JSON: '{}'".format(e),
            }

        super().__init__(error)
        self.response: Response = response
        self.error: Mapping[str, Any] = error
        self.code: int = self.error["code"]

    def __str__(self) -> str:
        return "{}: [{}]: {}".format(
            self.__class__.__name__, self.code, self.error["message"]
        )

    def __repr__(self) -> str:
        return self.__str__()

    def __reduce__(self) -> tuple:
        return self.__class__, (self.response,)


class SpreadsheetNotFound(GSpreadException):
    """Trying to open non-existent or inaccessible spreadsheet."""


class MissingIDsError(InvalidInputValue):
    """A cell inside the ID range is blank.

    Raised by
    :func:`gspread.worksheet.Worksheet.plan_row_reorder` before any plan is
    built, so that an incomplete data region is never silently reordered.

    :ivar blank_rows: 1-based row numbers of the blank cells.
    """

    def __init__(self, blank_rows: Any):
        self.blank_rows = list(blank_rows)
        super().__init__(
            "the ID range contains blank cells at row(s) {}".format(
                ", ".join(str(row) for row in self.blank_rows)
            )
        )


class DuplicateIDsError(InvalidInputValue):
    """An ID occurs more than once, so the order is not a permutation.

    :ivar duplicates: The duplicated IDs, sorted.
    :ivar source: Either ``"worksheet"`` or ``"target_order"``.
    """

    def __init__(self, duplicates: Any, source: str = "worksheet"):
        self.duplicates = list(duplicates)
        self.source = source
        super().__init__(
            "{} contains duplicate ID(s): {}".format(
                "the worksheet" if source == "worksheet" else "target_order",
                ", ".join(repr(id_) for id_ in self.duplicates),
            )
        )


class IDSetMismatchError(InvalidInputValue):
    """The worksheet IDs and the target order are not the same set.

    :ivar only_in_worksheet: IDs present in the range but missing from the
        target order.
    :ivar only_in_target: IDs present in the target order but missing from
        the range.
    """

    def __init__(self, only_in_worksheet: Any, only_in_target: Any):
        self.only_in_worksheet = list(only_in_worksheet)
        self.only_in_target = list(only_in_target)

        parts = []
        if self.only_in_worksheet:
            parts.append(
                "only in the worksheet: {}".format(
                    ", ".join(repr(id_) for id_ in self.only_in_worksheet)
                )
            )
        if self.only_in_target:
            parts.append(
                "only in target_order: {}".format(
                    ", ".join(repr(id_) for id_ in self.only_in_target)
                )
            )
        super().__init__(
            "the target order is not a permutation of the worksheet IDs ({})".format(
                "; ".join(parts)
            )
        )
