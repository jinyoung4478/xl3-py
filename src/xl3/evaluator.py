"""Expression evaluation given a row context.

Pure: takes an `Expr` and a `Context` and returns a value. Knows nothing
about openpyxl or template structure.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

from .errors import XtlError, xtl_error
from .expression import (
    BinOp,
    BoolLit,
    BracketRef,
    Expr,
    FuncCall,
    NumberLit,
    SourceRef,
    StringLit,
    UnaryNeg,
)
from .functions import get_simple_function
from .value_model import (
    canonical_string,
    compare_values,
    is_empty,
    is_truthy,
    parse_number_strict,
)


@dataclass
class EvalContext:
    """Per-evaluation context.

    `active_row` is the current data row's column→value mapping.
    `active_source_name` is the name of the source driving the active block
    (`"default"` unless `@source` overrides). It controls how
    `Source[Column]` references are resolved at row-level.
    `joined_rows` maps source-name → that source's current paired row inside
    the active block (populated by `@join`).
    `inputs` is the resolved `__inputs__[name]` lookup.
    `config_values` is the author-defined `__config__[key]` lookup.
    `active_row_set` is the filtered/sorted/topped row list for the active
    block — used by aggregates that take a bare `[Column]` argument.
    `named_sources` exposes every source's full row set + headers, used by
    cross-source aggregates and XLOOKUP.
    `row_index` is the 1-based index of the current data row inside its
    repeat block; None when ROW() is invalid (outside a repeat block).
    """

    active_row: dict[str, Any] = field(default_factory=dict)
    active_source_name: str = "default"
    active_source_columns: set[str] | None = None
    joined_rows: dict[str, dict[str, Any]] = field(default_factory=dict)
    joined_columns: dict[str, set[str]] = field(default_factory=dict)
    inputs: dict[str, Any] = field(default_factory=dict)
    config_values: dict[str, Any] = field(default_factory=dict)
    active_row_set: list[dict[str, Any]] | None = None
    named_sources: dict[str, dict[str, Any]] = field(default_factory=dict)
    row_index: int | None = None  # for ROW(); None outside a repeat block

    def lookup_active_column(self, name: str) -> Any:
        if self.active_source_columns is not None and name not in self.active_source_columns:
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{name}" does not exist in the active source',
            )
        return self.active_row.get(name)


def evaluate(expr: Expr, ctx: EvalContext) -> Any:
    """Evaluate `expr` under `ctx`. Returns a Python-native value."""
    if isinstance(expr, NumberLit):
        return expr.value
    if isinstance(expr, StringLit):
        return expr.value
    if isinstance(expr, BoolLit):
        return expr.value
    if isinstance(expr, BracketRef):
        return ctx.lookup_active_column(expr.column)
    if isinstance(expr, SourceRef):
        return _eval_source_ref(expr, ctx)
    if isinstance(expr, UnaryNeg):
        v = evaluate(expr.operand, ctx)
        return -_to_number(v)
    if isinstance(expr, BinOp):
        return _eval_binop(expr, ctx)
    if isinstance(expr, FuncCall):
        return _eval_call(expr, ctx)
    raise xtl_error("xl3/cell/numfmt-coercion", f"unhandled expression node: {expr!r}")


def _eval_source_ref(expr: SourceRef, ctx: EvalContext) -> Any:
    if expr.source == "__inputs__":
        return ctx.inputs.get(expr.column)
    if expr.source == "__config__":
        return ctx.config_values.get(expr.column)
    # `<active_source>[Column]` resolves to the active row's column.
    if expr.source == ctx.active_source_name:
        return ctx.lookup_active_column(expr.column)
    # `<joined_source>[Column]` resolves inside an @join block.
    if expr.source in ctx.joined_rows:
        cols = ctx.joined_columns.get(expr.source)
        if cols is not None and expr.column not in cols:
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{expr.column}" does not exist in source "{expr.source}"',
            )
        return ctx.joined_rows[expr.source].get(expr.column)
    raise xtl_error(
        "xl3/source/row-cross-block",
        f"Cannot reference {expr.source}[{expr.column}] outside an active "
        f"@source {expr.source} block",
    )


def _to_number(v: Any) -> float:
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    n = parse_number_strict(v)
    if n is None:
        # Excel coerces non-numeric to 0 in arithmetic; XTL doesn't pin this
        # explicitly, so we follow ECMA Number() coercion: fail → NaN, which
        # we surface via empty/non-finite handling.
        return float("nan")
    return n


def _eval_binop(expr: BinOp, ctx: EvalContext) -> Any:
    if expr.op == "&":
        left = canonical_string(evaluate(expr.left, ctx))
        right = canonical_string(evaluate(expr.right, ctx))
        return left + right
    if expr.op in {"+", "-", "*", "/"}:
        a = _to_number(evaluate(expr.left, ctx))
        b = _to_number(evaluate(expr.right, ctx))
        if expr.op == "+":
            return a + b
        if expr.op == "-":
            return a - b
        if expr.op == "*":
            return a * b
        if b == 0:
            return float("nan")  # division by zero — flow as empty per ADR-0009
        return a / b
    if expr.op in {"=", "!=", ">", "<", ">=", "<="}:
        a = evaluate(expr.left, ctx)
        b = evaluate(expr.right, ctx)
        c = compare_values(a, b)
        return {
            "=": c == 0,
            "!=": c != 0,
            ">": c > 0,
            "<": c < 0,
            ">=": c >= 0,
            "<=": c <= 0,
        }[expr.op]
    raise xtl_error("xl3/cell/numfmt-coercion", f"unknown operator {expr.op!r}")


def _eval_call(expr: FuncCall, ctx: EvalContext) -> Any:
    name = expr.name.upper()
    # ROW() — needs the repeat-block index from ctx.
    if name == "ROW":
        if len(expr.args) != 0:
            raise xtl_error(
                "xl3/eval/arity-mismatch",
                f"ROW: expected 0 arguments, got {len(expr.args)}",
            )
        if ctx.row_index is None:
            raise xtl_error(
                "xl3/cell/row-outside-repeat",
                "ROW() called outside a repeat block",
            )
        return ctx.row_index
    # IFS/IFERROR are lazy so unmatched branches and fallbacks are not evaluated.
    if name == "IFS":
        if len(expr.args) % 2 != 0:
            raise xtl_error(
                "xl3/eval/arity-mismatch",
                f"IFS: expected an even number of arguments, got {len(expr.args)}",
            )
        for i in range(0, len(expr.args), 2):
            if is_truthy(evaluate(expr.args[i], ctx)):
                return evaluate(expr.args[i + 1], ctx)
        raise xtl_error("xl3/eval/no-match", "IFS() no condition matched")
    if name == "IFERROR":
        if len(expr.args) != 2:
            raise xtl_error(
                "xl3/eval/arity-mismatch",
                f"IFERROR: expected 2 arguments, got {len(expr.args)}",
            )
        try:
            value = evaluate(expr.args[0], ctx)
        except XtlError:
            return evaluate(expr.args[1], ctx)
        if isinstance(value, float) and math.isnan(value):
            return evaluate(expr.args[1], ctx)
        return value
    # Aggregates — first arg is either a BracketRef (active row set) or a
    # SourceRef (named source's full row set). We DON'T evaluate the arg
    # eagerly; we inspect the AST to know which row set to fold over.
    if name in {"SUM", "COUNT", "AVERAGE", "AVG", "MIN", "MAX"}:
        return _eval_aggregate(name, expr.args, ctx)
    # XLOOKUP — needs source row sets.
    if name == "XLOOKUP":
        return _eval_xlookup(expr.args, ctx)
    # Context-free path
    fn = get_simple_function(name)
    if fn is None:
        raise xtl_error("xl3/cell/numfmt-coercion", f"unknown function {expr.name}()")
    args = [evaluate(a, ctx) for a in expr.args]
    return fn(args)


# ---------------------------------------------------------------------------
# Aggregates (ADR-0012: bare `[Col]` over active row set; `Source[Col]` over
# named source's full row set)
# ---------------------------------------------------------------------------


def _resolve_aggregate_target(
    arg: Any,
    ctx: EvalContext,
) -> tuple[list[dict[str, Any]] | None, str | None, str | None]:
    """Return (row_set, column, source_name).

    `row_set` is the list of rows over which to aggregate; None if the
    argument is invalid for aggregation. `column` is the column name to
    extract from each row; None for COUNT() (no argument). `source_name`
    is the source the row set came from (for unknown-column diagnostics).
    """
    if isinstance(arg, BracketRef):
        if ctx.active_row_set is None:
            return None, arg.column, ctx.active_source_name
        if (
            ctx.active_source_columns is not None
            and arg.column not in ctx.active_source_columns
        ):
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{arg.column}" does not exist in the active source',
            )
        return ctx.active_row_set, arg.column, ctx.active_source_name
    if isinstance(arg, SourceRef):
        if arg.source not in ctx.named_sources:
            raise xtl_error(
                "xl3/source/undeclared",
                f'Source "{arg.source}" is not declared in __sources__',
            )
        ns = ctx.named_sources[arg.source]
        headers = ns.get("headers", [])
        if arg.column not in headers:
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{arg.column}" does not exist in source "{arg.source}"',
            )
        return ns.get("rows", []), arg.column, arg.source
    return None, None, None


def _eval_aggregate(name: str, args: list[Any], ctx: EvalContext) -> Any:
    if name == "COUNT" and len(args) == 0:
        # COUNT() — count rows in the active row set.
        if ctx.active_row_set is None:
            return 0
        return len(ctx.active_row_set)
    if len(args) != 1:
        raise xtl_error(
            "xl3/eval/arity-mismatch",
            f"{name}: expected 1 argument, got {len(args)}",
        )
    arg_ast = args[0]
    rows, column, _src = _resolve_aggregate_target(arg_ast, ctx)
    if rows is None or column is None:
        raise xtl_error(
            "xl3/cell/numfmt-coercion",
            f"{name} argument must be a column reference",
        )
    values = [r.get(column) for r in rows]
    if name == "COUNT":
        # COUNT([col]) — count non-empty values per ADR-0007.
        from .value_model import is_empty as _is_empty

        return sum(1 for v in values if not _is_empty(v))
    # Numeric aggregates — coerce values, skipping empty / unparseable ones.
    from .value_model import is_empty as _is_empty

    nums: list[float] = []
    for v in values:
        if _is_empty(v):
            continue
        if isinstance(v, bool):
            nums.append(1.0 if v else 0.0)
        elif isinstance(v, (int, float)):
            nums.append(float(v))
        else:
            from .value_model import parse_number_strict as _pn

            n = _pn(v)
            if n is not None:
                nums.append(n)
    if name == "SUM":
        return sum(nums) if nums else 0
    if name in ("AVERAGE", "AVG"):
        return sum(nums) / len(nums) if nums else 0
    if name == "MIN":
        # MIN/MAX: spec says aggregates operate on the row set; for date columns,
        # min/max should compare by underlying timestamp. Use original values
        # for non-numeric (date) min/max via compare_values.
        return _aggregate_extremum(values, "min")
    if name == "MAX":
        return _aggregate_extremum(values, "max")
    raise xtl_error("xl3/cell/numfmt-coercion", f"unknown aggregate {name}")


def _aggregate_extremum(values: list[Any], kind: str) -> Any:
    from .value_model import compare_values as _cmp
    from .value_model import is_empty as _is_empty

    non_empty = [v for v in values if not _is_empty(v)]
    if not non_empty:
        return None
    best = non_empty[0]
    for v in non_empty[1:]:
        c = _cmp(v, best)
        if (kind == "min" and c < 0) or (kind == "max" and c > 0):
            best = v
    return best


# ---------------------------------------------------------------------------
# XLOOKUP (ADR-0013) — exact match only, optional fallback
# ---------------------------------------------------------------------------


def _eval_xlookup(args: list[Any], ctx: EvalContext) -> Any:
    from .value_model import canonical_string as _cs
    from .value_model import compare_values as _cmp

    if len(args) not in (3, 4):
        raise xtl_error(
            "xl3/eval/arity-mismatch",
            f"XLOOKUP: expected 3 or 4 arguments, got {len(args)}",
        )
    lookup_value = evaluate(args[0], ctx)
    lookup_arg = args[1]
    return_arg = args[2]
    fallback_provided = len(args) == 4
    fallback_value = evaluate(args[3], ctx) if fallback_provided else None
    # Both arrays MUST be SourceRef with the same source (ADR-0013).
    if not isinstance(lookup_arg, SourceRef) or not isinstance(return_arg, SourceRef):
        raise xtl_error(
            "xl3/xlookup/bare-bracket",
            "XLOOKUP arg 2 and arg 3 must be a source-prefixed bracket reference",
        )
    if lookup_arg.source != return_arg.source:
        raise xtl_error(
            "xl3/xlookup/source-mismatch",
            "XLOOKUP arg 2 and arg 3 must match (same source)",
        )
    src_name = lookup_arg.source
    if src_name not in ctx.named_sources:
        raise xtl_error(
            "xl3/source/undeclared",
            f'Source "{src_name}" is not declared in __sources__',
        )
    src = ctx.named_sources[src_name]
    headers = src.get("headers", [])
    if lookup_arg.column not in headers:
        raise xtl_error(
            "xl3/source/unknown-column",
            f'Column "{lookup_arg.column}" does not exist in source "{src_name}"',
        )
    if return_arg.column not in headers:
        raise xtl_error(
            "xl3/source/unknown-column",
            f'Column "{return_arg.column}" does not exist in source "{src_name}"',
        )
    rows = src.get("rows", [])
    for r in rows:
        if _cmp(r.get(lookup_arg.column), lookup_value) == 0:
            return r.get(return_arg.column)
    if fallback_provided:
        return fallback_value
    raise xtl_error(
        "xl3/xlookup/no-match",
        f'XLOOKUP found no row where [{lookup_arg.column}] equals "{_cs(lookup_value)}"',
    )
