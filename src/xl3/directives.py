"""Directive parsing + row-set transform pipeline.

Directives are template expressions whose body starts with `@`:

    @filter [field] op value | [field] in __lists__[name] | [field] !in __lists__[name]
    @sort [field] [asc|desc]
    @top N
    @repeat [right] [N]
    @source SourceName
    @join JoinedSource on JoinedSource[col] = PrimarySource[col]

Per evaluation.md, directives apply: @source/@join → @filter → @sort → @top,
then the data row is expanded.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Literal

from .errors import xtl_error
from .expression import (
    BinOp,
    BracketRef,
    Expr,
    SourceRef,
    parse_expression,
)
from .value_model import canonical_string, compare_values, is_empty


# ---------------------------------------------------------------------------
# Directive AST
# ---------------------------------------------------------------------------


@dataclass
class FilterDirective:
    """`@filter [field] op value` or `@filter [field] in/!in __lists__[name]`."""

    column: str  # the active-source column being filtered
    op: str  # = != > < >= <= in !in
    rhs: Expr | None = None  # for comparison operators
    list_name: str | None = None  # for in / !in


@dataclass
class SortKey:
    column: str
    direction: Literal["asc", "desc"] = "asc"


@dataclass
class SortDirective:
    keys: list[SortKey]


@dataclass
class TopDirective:
    n: int


@dataclass
class RepeatRightDirective:
    col_span: int = 1


@dataclass
class SourceDirective:
    source_name: str


@dataclass
class JoinDirective:
    joined_source: str  # name of the source being joined in
    primary_source: str  # name of the active source (default or @source X)
    joined_column: str
    primary_column: str


Directive = (
    FilterDirective
    | SortDirective
    | TopDirective
    | RepeatRightDirective
    | SourceDirective
    | JoinDirective
)


# ---------------------------------------------------------------------------
# Directive parser
# ---------------------------------------------------------------------------


_DIRECTIVE_RE = re.compile(r"^\s*@(\w+)\b\s*(.*)$", re.DOTALL)


class DirectiveParseError(Exception):
    pass


def is_directive_body(body: str) -> bool:
    return body.lstrip().startswith("@")


def parse_directive(body: str) -> Directive:
    m = _DIRECTIVE_RE.match(body.strip())
    if not m:
        raise DirectiveParseError(f"not a directive: {body!r}")
    name = m.group(1).lower()
    rest = m.group(2).strip()
    if name == "filter":
        return _parse_filter(rest)
    if name == "sort":
        return _parse_sort(rest)
    if name == "top":
        return _parse_top(rest)
    if name == "repeat":
        return _parse_repeat(rest)
    if name == "source":
        return _parse_source(rest)
    if name == "join":
        return _parse_join(rest)
    raise DirectiveParseError(f"unknown directive @{name}")


_BRACKET_COL_RE = re.compile(r"^\[\s*([^\]]+?)\s*\]\s*(.*)$", re.DOTALL)
_LIST_REF_RE = re.compile(r"^__lists__\[\s*([^\]]+?)\s*\]\s*$", re.DOTALL)


def _parse_filter(rest: str) -> FilterDirective:
    """`[col] op value` or `[col] in __lists__[name]` (or `!in`)."""
    m = _BRACKET_COL_RE.match(rest)
    if not m:
        raise DirectiveParseError(
            f"@filter must start with [Column], got {rest!r}"
        )
    column = m.group(1).strip()
    after = m.group(2).strip()
    # Detect `in` / `!in` first
    op_match = re.match(r"^(!?in)\b\s*(.*)$", after, re.DOTALL)
    if op_match:
        op = op_match.group(1)
        list_text = op_match.group(2).strip()
        lm = _LIST_REF_RE.match(list_text)
        if not lm:
            raise DirectiveParseError(
                f"@filter ... {op} requires __lists__[name], got {list_text!r}"
            )
        return FilterDirective(column=column, op=op, list_name=lm.group(1).strip())
    # Comparison: `op value`
    cmp_match = re.match(r"^(!=|>=|<=|=|>|<)\s*(.+)$", after, re.DOTALL)
    if not cmp_match:
        raise DirectiveParseError(f"@filter operator missing in {after!r}")
    op = cmp_match.group(1)
    rhs_body = cmp_match.group(2).strip()
    try:
        rhs_expr = parse_expression(rhs_body)
    except Exception as e:  # noqa: BLE001
        raise DirectiveParseError(f"@filter rhs parse failed: {e}") from e
    return FilterDirective(column=column, op=op, rhs=rhs_expr)


def _parse_sort(rest: str) -> SortDirective:
    m = _BRACKET_COL_RE.match(rest)
    if not m:
        raise DirectiveParseError(f"@sort must start with [Column], got {rest!r}")
    column = m.group(1).strip()
    direction_text = m.group(2).strip().lower()
    if direction_text == "" or direction_text == "asc":
        direction: Literal["asc", "desc"] = "asc"
    elif direction_text == "desc":
        direction = "desc"
    else:
        raise DirectiveParseError(f"@sort direction must be asc or desc, got {direction_text!r}")
    return SortDirective(keys=[SortKey(column=column, direction=direction)])


def _parse_top(rest: str) -> TopDirective:
    try:
        n = int(rest.strip())
    except ValueError as e:
        raise DirectiveParseError(f"@top N must be an integer, got {rest!r}") from e
    if n < 0:
        raise DirectiveParseError(f"@top N must be non-negative, got {n}")
    return TopDirective(n=n)


def _parse_repeat(rest: str) -> RepeatRightDirective:
    parts = rest.split()
    if not parts:
        raise DirectiveParseError("@repeat must specify direction (right)")
    direction = parts[0].lower()
    if direction != "right":
        raise DirectiveParseError(f"@repeat direction must be 'right', got {direction!r}")
    if len(parts) == 1:
        return RepeatRightDirective(col_span=1)
    try:
        span = int(parts[1])
    except ValueError as e:
        raise DirectiveParseError(f"@repeat right N must be an integer, got {parts[1]!r}") from e
    if span < 1:
        raise DirectiveParseError(f"@repeat right N must be >= 1, got {span}")
    return RepeatRightDirective(col_span=span)


def _parse_source(rest: str) -> SourceDirective:
    name = rest.strip()
    if not name or not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name):
        raise DirectiveParseError(f"@source name must be a valid identifier, got {name!r}")
    return SourceDirective(source_name=name)


_JOIN_RE = re.compile(
    r"^([A-Za-z_][A-Za-z0-9_]*)\s+on\s+(.+)$",
    re.DOTALL,
)


def _parse_join(rest: str) -> JoinDirective:
    m = _JOIN_RE.match(rest.strip())
    if not m:
        raise DirectiveParseError(f"@join must be 'JoinedSource on <on-clause>', got {rest!r}")
    joined = m.group(1)
    on_body = m.group(2).strip()
    # The on-clause is `Source[col] = OtherSource[col]`. Parse it through the
    # expression parser then validate shape.
    try:
        on_expr = parse_expression(on_body)
    except Exception as e:  # noqa: BLE001
        raise DirectiveParseError(f"@join on-clause parse failed: {e}") from e
    if not isinstance(on_expr, BinOp) or on_expr.op != "=":
        raise DirectiveParseError(
            "@join on-clause must be '<JoinedSource>[col] = <PrimarySource>[col]'"
        )
    if not (isinstance(on_expr.left, SourceRef) and isinstance(on_expr.right, SourceRef)):
        raise DirectiveParseError(
            "@join key columns must reference the joined and primary sources"
        )
    left, right = on_expr.left, on_expr.right
    # One side is `joined`; the other is the primary source.
    if left.source == joined:
        joined_col, primary_col = left.column, right.column
        primary_source = right.source
    elif right.source == joined:
        joined_col, primary_col = right.column, left.column
        primary_source = left.source
    else:
        raise DirectiveParseError(
            "@join key columns must reference the joined and primary sources"
        )
    return JoinDirective(
        joined_source=joined,
        primary_source=primary_source,
        joined_column=joined_col,
        primary_column=primary_col,
    )


# ---------------------------------------------------------------------------
# Row-set transform pipeline
# ---------------------------------------------------------------------------


@dataclass
class BlockDirectives:
    """Collected directives for one data block, in declaration order."""

    source_directive: SourceDirective | None = None
    join_directive: JoinDirective | None = None
    filters: list[FilterDirective] = field(default_factory=list)
    sorts: list[SortDirective] = field(default_factory=list)
    top: TopDirective | None = None
    repeat_right: RepeatRightDirective | None = None
    raw: list[Directive] = field(default_factory=list)

    def add(self, d: Directive) -> None:
        self.raw.append(d)
        if isinstance(d, SourceDirective):
            self.source_directive = d
        elif isinstance(d, JoinDirective):
            self.join_directive = d
        elif isinstance(d, FilterDirective):
            self.filters.append(d)
        elif isinstance(d, SortDirective):
            self.sorts.append(d)
        elif isinstance(d, TopDirective):
            self.top = d
        elif isinstance(d, RepeatRightDirective):
            self.repeat_right = d


def apply_filters(
    rows: list[dict[str, Any]],
    filters: list[FilterDirective],
    list_sheets: dict[str, list[str]],
) -> list[dict[str, Any]]:
    out = list(rows)
    for fd in filters:
        out = [r for r in out if _row_passes(r, fd, list_sheets)]
    return out


def _row_passes(
    row: dict[str, Any],
    fd: FilterDirective,
    list_sheets: dict[str, list[str]],
) -> bool:
    val = row.get(fd.column)
    if fd.op in ("in", "!in"):
        if fd.list_name is None:
            return False
        if fd.list_name not in list_sheets:
            raise xtl_error(
                "xl3/lists/missing-reference",
                f'List "{fd.list_name}" is not declared in __lists__',
            )
        # Empty values never match `in` (ADR-0007 §"List Sheets")
        if is_empty(val):
            return fd.op == "!in"
        wanted = canonical_string(val)
        contains = wanted in list_sheets[fd.list_name]
        return contains if fd.op == "in" else not contains
    # Comparison: rhs already an Expr — evaluate as a literal-only Expr (no
    # source-column refs allowed in @filter rhs at this scope).
    from .evaluator import EvalContext, evaluate

    rhs_val = evaluate(fd.rhs, EvalContext()) if fd.rhs is not None else None
    c = compare_values(val, rhs_val)
    return {
        "=": c == 0,
        "!=": c != 0,
        ">": c > 0,
        "<": c < 0,
        ">=": c >= 0,
        "<=": c <= 0,
    }[fd.op]


def apply_sorts(
    rows: list[dict[str, Any]],
    sorts: list[SortDirective],
) -> list[dict[str, Any]]:
    """Stable multi-key sort. First @sort = primary key, later = tiebreakers."""
    if not sorts:
        return rows
    # Flatten directives into a single list of (column, direction) keys.
    keys: list[SortKey] = []
    for s in sorts:
        keys.extend(s.keys)
    # Apply LAST-to-first because Python's sort is stable.
    out = list(rows)
    for k in reversed(keys):
        reverse = k.direction == "desc"

        def cmp_key(r: dict[str, Any], col: str = k.column, rev: bool = reverse) -> Any:
            v = r.get(col)
            return _SortableKey(v, rev)

        out.sort(key=cmp_key)
    return out


class _SortableKey:
    """Wrapper that makes any value comparable via compare_values, with
    desc handling reversed (so a single Python sort can do mixed directions).

    Note: when all sort directives share a direction, callers can use
    `reverse=True` instead. Here we keep the wrapper so future mixed-direction
    multi-key sorts work uniformly.
    """

    __slots__ = ("v", "reverse")

    def __init__(self, v: Any, reverse: bool) -> None:
        self.v = v
        self.reverse = reverse

    def __lt__(self, other: "_SortableKey") -> bool:
        c = compare_values(self.v, other.v)
        if self.reverse:
            c = -c
        return c < 0


def apply_top(rows: list[dict[str, Any]], top: TopDirective | None) -> list[dict[str, Any]]:
    if top is None:
        return rows
    return rows[: top.n]
