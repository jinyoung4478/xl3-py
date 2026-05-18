"""Stable error code catalog for xl3 (ADR-0015).

Every spec-defined error carries a stable `code` of the form `xl3/<category>/<id>`.
Hosts use the code for localization and programmatic dispatch; the English
`Error.message` (here, `XtlError.args[0]`) remains the conformance contract.
"""

from __future__ import annotations

from typing import Final, Literal, get_args

XtlErrorCode = Literal[
    # Config
    "xl3/config/source-table-removed",
    "xl3/config/invalid-source-table",
    # Inputs (ADR-0010)
    "xl3/inputs/missing-required",
    "xl3/inputs/parse-number",
    "xl3/inputs/parse-date",
    "xl3/inputs/select-option",
    "xl3/inputs/duplicate-name",
    "xl3/inputs/invalid-name",
    "xl3/inputs/invalid-type",
    "xl3/inputs/conflict-config",
    "xl3/inputs/missing-header",
    "xl3/inputs/missing-options",
    "xl3/inputs/forward-reference",
    "xl3/inputs/runtime-only-fn",
    # Sources (ADR-0012)
    "xl3/source/undeclared",
    "xl3/source/sheet-missing",
    "xl3/source/duplicate-name",
    "xl3/source/invalid-name",
    "xl3/source/missing-header",
    "xl3/source/missing-required",
    "xl3/source/row-cross-block",
    "xl3/source/unknown-column",
    "xl3/sources/not-a-dictionary",
    # Grouping / subtotal rows (ADR-0038)
    "xl3/group/missing-key",
    "xl3/subtotal/outside-group",
    "xl3/subtotal/bad-aggregate",
    # Reserved sheets (ADR-0011)
    "xl3/sheet/reserved-name",
    "xl3/sheet/duplicate-list-name",
    # Join (ADR-0014)
    "xl3/join/undeclared-source",
    "xl3/join/bad-on-clause",
    # Lists
    "xl3/lists/missing-reference",
    # Cell evaluation
    "xl3/cell/numfmt-coercion",
    "xl3/cell/row-outside-repeat",
    "xl3/cell/formula-no-cache",
    # Evaluator (ADR-0044)
    "xl3/eval/type-mismatch",
    "xl3/eval/arity-mismatch",
    "xl3/eval/no-match",
    # XLOOKUP (ADR-0013)
    "xl3/xlookup/no-match",
    "xl3/xlookup/source-mismatch",
    "xl3/xlookup/bare-bracket",
    # Filename sanitization (ADR-0002)
    "xl3/filename/empty",
    "xl3/filename/too-long",
]


_ALL_CODES: Final[frozenset[str]] = frozenset(get_args(XtlErrorCode))


class XtlError(Exception):
    """Spec-defined xl3 error with a stable `code`."""

    code: str

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code

    def __str__(self) -> str:
        return self.args[0] if self.args else ""


def xtl_error(code: XtlErrorCode, message: str) -> XtlError:
    """Construct an XtlError. The `code` MUST appear in the catalog above."""
    if code not in _ALL_CODES:
        raise ValueError(f"unknown XtlErrorCode: {code!r}")
    return XtlError(code, message)


def is_xtl_error(e: object) -> bool:
    """True iff `e` is an XtlError carrying a recognized `xl3/...` code."""
    return (
        isinstance(e, XtlError)
        and isinstance(getattr(e, "code", None), str)
        and e.code.startswith("xl3/")
    )
