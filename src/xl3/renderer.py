"""Render output workbook(s) from parsed template + sources.

Operates on the block plan produced by `parser.py`:
  - StaticRowPlan: emit one row, expressions evaluated with no active row.
  - DataRowPlan:   apply directives → expand once per (filtered/sorted/top) source row.

Supports:
  - @source SourceName       (block iterates the named source)
  - @join JoinedSource on …  (inner-join: pair primary with first match)
  - @filter / @sort / @top   (transform the block's row set)
  - @repeat right [N]        (horizontal expansion)

NOT yet:
  - Multi-file groups (output_file_pattern with bare-ident group keys)
  - Sheet-name group keys
  - numFmt-driven coercion (ADR-0003)
  - Filename sanitization (ADR-0002)
"""

from __future__ import annotations

from copy import copy
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell

from .directives import (
    BlockDirectives,
    JoinDirective,
    apply_filters,
    apply_sorts,
    apply_top,
)
from .errors import xtl_error
from .evaluator import EvalContext, evaluate
from .expression import (
    CellTemplate,
    DirectiveSegment,
    ExprSegment,
    TextSegment,
)
from .parser import (
    DataRowPlan,
    ParsedTemplate,
    SheetTemplate,
    StaticRowPlan,
    TemplateCell,
    is_reserved_sheet,
)
from .reader import SourceData
from .types import OutputFile
from .value_model import canonical_string


def render(
    parsed: ParsedTemplate,
    sources: dict[str, SourceData],
    template_bytes: bytes,
    config_values: dict[str, Any] | None = None,
    inputs: dict[str, Any] | None = None,
) -> list[OutputFile]:
    """Render output files. ADR-0016 ordering: file groups by first-seen,
    sheet groups within a file by first-seen."""
    config_values = config_values or {}
    inputs = inputs or {}

    default = sources.get("default")
    default_rows = list(default.rows) if default else []

    file_group_keys = _extract_file_group_keys(parsed.meta.output_file_pattern or "")
    file_groups = _partition_first_seen(default_rows, file_group_keys)
    if not file_groups:
        # No source rows at all — emit a single file with empty data block
        # expansions (preserves headers / static content). Fixture 031 covers
        # the explicit zero-row case via `_should_suppress_output`.
        file_groups = [_GroupBucket(key=(), rows=[])]

    files: list[OutputFile] = []
    for bucket in file_groups:
        # Build a per-file sources dict with the default's rows narrowed to
        # this file's bucket. Named sources keep their full row sets.
        file_sources = dict(sources)
        if default is not None:
            file_sources = dict(sources)
            file_sources["default"] = SourceData(
                sheet_name=default.sheet_name,
                headers=list(default.headers),
                rows=list(bucket.rows),
            )

        wb = load_workbook(BytesIO(template_bytes))

        # Render each non-reserved sheet — splitting that sheet by sheet-group
        # keys when its name is a group-key template like `{{ Region }}`.
        for st in parsed.sheets:
            sheet_group_keys = _extract_sheet_group_keys(st.original_name)
            if sheet_group_keys:
                _render_grouped_sheet(
                    wb, st, sheet_group_keys, file_sources, config_values, inputs, parsed
                )
            else:
                ws = wb[st.original_name]
                _render_sheet(ws, st, file_sources, config_values, inputs, parsed)

        for sn in list(wb.sheetnames):
            if is_reserved_sheet(sn):
                del wb[sn]

        if _should_suppress_output(parsed, file_sources):
            continue

        out_io = BytesIO()
        wb.save(out_io)

        filename = _evaluate_filename(parsed, file_sources, config_values, inputs)
        sanitized, _warning = _sanitize_via_filename_module(filename)
        files.append(OutputFile(filename=sanitized, data=out_io.getvalue()))
    return files


# ---------------------------------------------------------------------------
# Group-splitting helpers (ADR-0016)
# ---------------------------------------------------------------------------


from dataclasses import dataclass as _dc


@_dc
class _GroupBucket:
    key: tuple[Any, ...]
    rows: list[dict[str, Any]]


_BARE_IDENT_BLOCK_RE = __import__("re").compile(
    r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}"
)


def _extract_file_group_keys(pattern: str) -> list[str]:
    """Bare-identifier `{{ X }}` blocks in the output_file_pattern are file
    group keys. Bracketed `{{ [Col] }}` is a row-substitution, not a group
    key — those expressions evaluate against the first row of each file
    group, not partition the rows.
    """
    return [m.group(1) for m in _BARE_IDENT_BLOCK_RE.finditer(pattern)]


def _extract_sheet_group_keys(sheet_name: str) -> list[str]:
    """Bare-identifier `{{ X }}` blocks in a sheet name are sheet group keys."""
    return [m.group(1) for m in _BARE_IDENT_BLOCK_RE.finditer(sheet_name)]


def _partition_first_seen(
    rows: list[dict[str, Any]], group_keys: list[str]
) -> list[_GroupBucket]:
    """Walk `rows` once; group by the tuple of `group_keys` values.

    Buckets are returned in **first-seen** order (per ADR-0016): the first
    row whose key is X causes bucket X to be emitted first.
    """
    if not group_keys:
        return [_GroupBucket(key=(), rows=list(rows))] if rows else []
    seen: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    order: list[tuple[Any, ...]] = []
    for r in rows:
        key = tuple(canonical_string(r.get(k)) for k in group_keys)
        if key not in seen:
            seen[key] = []
            order.append(key)
        seen[key].append(r)
    return [_GroupBucket(key=k, rows=seen[k]) for k in order]


def _render_grouped_sheet(
    wb: Any,
    st: SheetTemplate,
    sheet_group_keys: list[str],
    sources: dict[str, SourceData],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
    parsed: ParsedTemplate,
) -> None:
    """Sheet-name template like `{{ Region }}`: clone the template sheet
    once per distinct sheet-group key combination (first-seen order) and
    render each with its bucket of rows.
    """
    default = sources.get("default")
    rows = list(default.rows) if default else []
    buckets = _partition_first_seen(rows, sheet_group_keys)
    if not buckets:
        # No source rows hit this file group — drop the template sheet entirely.
        if st.original_name in wb.sheetnames:
            del wb[st.original_name]
        return
    template_idx = wb.sheetnames.index(st.original_name)
    template_ws = wb[st.original_name]
    new_sheets: list[tuple[str, _GroupBucket]] = []
    for bucket in buckets:
        new_name = canonical_string(bucket.key[0]) if len(bucket.key) == 1 else "_".join(
            canonical_string(v) for v in bucket.key
        )
        if not new_name:
            new_name = st.original_name
        # Copy the template sheet using openpyxl's built-in copy
        copied = wb.copy_worksheet(template_ws)
        copied.title = new_name
        new_sheets.append((new_name, bucket))
    # Now remove the template sheet (it's been copied N times)
    del wb[st.original_name]
    # Move the copies to the template's original position
    for i, (name, bucket) in enumerate(new_sheets):
        ws = wb[name]
        # Restrict default source rows for this sheet's render
        per_sheet_sources = dict(sources)
        if default is not None:
            per_sheet_sources["default"] = SourceData(
                sheet_name=default.sheet_name,
                headers=list(default.headers),
                rows=list(bucket.rows),
            )
        _render_sheet(ws, st, per_sheet_sources, config_values, inputs, parsed)
        # Place the rendered sheet at the original index (keeps order tight)
        wb.move_sheet(ws, offset=template_idx + i - wb.sheetnames.index(name))


def _should_suppress_output(
    parsed: ParsedTemplate, sources: dict[str, SourceData]
) -> bool:
    """Fixture 031: when the default source has zero rows, produce no output.

    Per the ADR-0008 deferred gap, fixture 031 freezes the answer for the
    explicit zero-data `source_table` range case: no output workbook.
    """
    default = sources.get("default")
    return default is not None and not default.rows


def _sanitize_via_filename_module(filename: str) -> tuple[str, str | None]:
    from .filename import sanitize_filename

    return sanitize_filename(filename)


def _rename_group_sheets(
    wb: Any,
    parsed: ParsedTemplate,
    sources: dict[str, SourceData],
) -> None:
    """Apply `{{ <ident> }}` sheet-name templating using the first source row.

    Multi-group splitting (one rendered sheet per distinct group key) is the
    right answer per ADR-0016, but for the bootstrap we cover the single-row
    or single-group case which is what fixtures 086 needs.
    """
    import re as _re

    default_source = sources.get("default")
    if not default_source or not default_source.rows:
        return
    pattern = _re.compile(r"^\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}$")
    for sn in list(wb.sheetnames):
        m = pattern.match(sn)
        if not m:
            continue
        col = m.group(1)
        first_val = default_source.rows[0].get(col)
        if first_val is None:
            continue
        new_name = canonical_string(first_val)
        if new_name and new_name != sn:
            wb[sn].title = new_name


def _render_sheet(
    ws: Any,
    st: SheetTemplate,
    sources: dict[str, SourceData],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
    parsed: ParsedTemplate,
) -> None:
    # Cache cell styles before mutating the worksheet.
    style_cache = _capture_styles(ws, st)

    # Compute the row range we need to clear: all rows referenced by the plan
    # PLUS any directive-only rows. Directive rows must be stripped from the
    # output; if we leave them populated we'd emit literal `{{ @repeat right }}`
    # text (fixture 004 regression).
    template_rows_used: set[int] = set()
    for plan in st.plan:
        template_rows_used.add(plan.template_row)
    template_rows_used |= st.directive_only_rows
    if template_rows_used:
        min_r = min(template_rows_used)
        max_r = ws.max_row or max(template_rows_used)
        for r in range(min_r, max_r + 1):
            for c in range(1, st.max_col + 1):
                ws.cell(row=r, column=c).value = None

    out_row = min(template_rows_used) if template_rows_used else 1
    for plan in st.plan:
        if isinstance(plan, StaticRowPlan):
            _emit_static(ws, plan, out_row, style_cache, sources, config_values, inputs)
            out_row += 1
        else:
            out_row = _emit_data_block(
                ws,
                plan,
                out_row,
                sources,
                style_cache,
                config_values,
                inputs,
            )


def _capture_styles(
    ws: Any, st: SheetTemplate
) -> dict[tuple[int, int], Any]:
    cache: dict[tuple[int, int], Any] = {}
    for plan in st.plan:
        for c in range(1, st.max_col + 1):
            cell = ws.cell(row=plan.template_row, column=c)
            cache[(plan.template_row, c)] = (
                copy(cell.font),
                copy(cell.fill),
                copy(cell.border),
                copy(cell.alignment),
                copy(cell.number_format),
            )
    return cache


def _apply_style(cell: Cell, style: Any) -> None:
    if style is None:
        return
    font, fill, border, align, fmt = style
    cell.font = font
    cell.fill = fill
    cell.border = border
    cell.alignment = align
    cell.number_format = fmt


def _emit_static(
    ws: Any,
    plan: StaticRowPlan,
    out_row: int,
    style_cache: dict[tuple[int, int], Any],
    sources: dict[str, SourceData],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
) -> None:
    # Static cells can still reference cross-source aggregates and XLOOKUP
    # over named sources, so expose them on the context. The default
    # source's full row set serves as `active_row_set` for bare-bracket
    # aggregates that appear in static cells (e.g., a totals row).
    default_source = sources.get("default")
    ctx = EvalContext(
        active_row={},
        inputs=inputs,
        config_values=config_values,
        active_source_columns=None,
        named_sources=_build_named_sources_view(sources),
        active_row_set=list(default_source.rows) if default_source else None,
    )
    for tc in plan.cells:
        value = _render_cell(tc.template, ctx)
        style = style_cache.get((plan.template_row, tc.col))
        if tc.template.is_single_expression and style is not None:
            value = _apply_numfmt_coercion(value, style[4])
        target = ws.cell(row=out_row, column=tc.col, value=value)
        _apply_style(target, style)


def _emit_data_block(
    ws: Any,
    plan: DataRowPlan,
    out_row: int,
    sources: dict[str, SourceData],
    style_cache: dict[tuple[int, int], Any],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
) -> int:
    """Emit one expanded data block; return the next free output row."""
    bd = plan.directives

    # Resolve the active source.
    active_source_name = bd.source_directive.source_name if bd.source_directive else "default"
    if active_source_name not in sources:
        raise xtl_error(
            "xl3/source/undeclared",
            f'Source "{active_source_name}" is not declared in __sources__',
        )
    primary = sources[active_source_name]

    # Apply filter / sort / top to the primary row set.
    rows = list(primary.rows)
    rows = apply_filters(rows, bd.filters, _collect_lists(plan, sources, config_values, inputs))
    rows = apply_sorts(rows, bd.sorts)
    rows = apply_top(rows, bd.top)

    # Resolve the join, if any.
    join: JoinDirective | None = bd.join_directive
    joined_rows_for_primary: list[dict[str, Any] | None] = [None] * len(rows)
    if join is not None:
        if join.joined_source not in sources:
            raise xtl_error(
                "xl3/join/undeclared-source",
                f'@join source "{join.joined_source}" must be declared in __sources__',
            )
        # The "other side" of the on-clause must name the block's active
        # source — otherwise the join can't be satisfied.
        if join.primary_source != active_source_name:
            raise xtl_error(
                "xl3/join/bad-on-clause",
                "@join key columns must reference the joined and primary sources",
            )
        joined_data = sources[join.joined_source]
        # Validate columns referenced by the on-clause exist on each side.
        if join.primary_column not in primary.headers:
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{join.primary_column}" does not exist in source "{join.primary_source}"',
            )
        if join.joined_column not in joined_data.headers:
            raise xtl_error(
                "xl3/source/unknown-column",
                f'Column "{join.joined_column}" does not exist in source "{join.joined_source}"',
            )
        kept: list[dict[str, Any]] = []
        kept_pairs: list[dict[str, Any]] = []
        for r in rows:
            primary_key = r.get(join.primary_column)
            match = _first_match(joined_data.rows, join.joined_column, primary_key)
            if match is None:
                continue  # inner-join semantics drop unmatched rows
            kept.append(r)
            kept_pairs.append(match)
        rows = kept
        joined_rows_for_primary = list(kept_pairs)  # type: ignore[assignment]

    # Repeat-right vs default vertical expansion.
    if bd.repeat_right is not None:
        return _emit_repeat_right(
            ws,
            plan,
            out_row,
            rows,
            joined_rows_for_primary,
            join,
            sources,
            active_source_name,
            primary,
            style_cache,
            config_values,
            inputs,
            bd.repeat_right.col_span,
        )
    return _emit_vertical(
        ws,
        plan,
        out_row,
        rows,
        joined_rows_for_primary,
        join,
        sources,
        active_source_name,
        primary,
        style_cache,
        config_values,
        inputs,
    )


def _collect_lists(
    plan: DataRowPlan,
    sources: dict[str, SourceData],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
) -> dict[str, list[str]]:
    # Lists live on the parsed template, not on per-block data; resolve from
    # the parser's result via a closure when we wire it in. Caller passes the
    # template in render(), but apply_filters needs lookup_lists at runtime.
    # We instead expose `lists` via plan? Cleaner: expose via a module-level
    # holder. Use a thread-local-style fallback since each render() call is
    # synchronous.
    return _RENDER_LISTS.get() or {}


# A simple holder so _emit_data_block/_collect_lists can see the parsed
# template's list_sheets without threading the dict through every call.
class _ListHolder:
    _stack: list[dict[str, list[str]]] = []

    def push(self, lists: dict[str, list[str]]) -> None:
        self._stack.append(lists)

    def pop(self) -> None:
        self._stack.pop()

    def get(self) -> dict[str, list[str]] | None:
        return self._stack[-1] if self._stack else None


_RENDER_LISTS = _ListHolder()


def _first_match(
    joined_rows: list[dict[str, Any]],
    joined_column: str,
    primary_key: Any,
) -> dict[str, Any] | None:
    from .value_model import compare_values

    for r in joined_rows:
        if compare_values(r.get(joined_column), primary_key) == 0:
            return r
    return None


def _emit_vertical(
    ws: Any,
    plan: DataRowPlan,
    out_row: int,
    rows: list[dict[str, Any]],
    joined_rows: list[dict[str, Any] | None],
    join: JoinDirective | None,
    sources: dict[str, SourceData],
    active_source_name: str,
    primary: SourceData,
    style_cache: dict[tuple[int, int], Any],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
) -> int:
    for i, src_row in enumerate(rows):
        ctx = _build_row_context(
            src_row,
            joined_rows[i],
            join,
            sources,
            active_source_name,
            primary,
            i + 1,
            config_values,
            inputs,
            rows,
        )
        for tc in plan.cells:
            value = _render_cell(tc.template, ctx)
            style = style_cache.get((plan.template_row, tc.col))
            if tc.template.is_single_expression and style is not None:
                value = _apply_numfmt_coercion(value, style[4])
            target = ws.cell(row=out_row, column=tc.col, value=value)
            _apply_style(target, style)
        out_row += 1
    return out_row


def _emit_repeat_right(
    ws: Any,
    plan: DataRowPlan,
    out_row: int,
    rows: list[dict[str, Any]],
    joined_rows: list[dict[str, Any] | None],
    join: JoinDirective | None,
    sources: dict[str, SourceData],
    active_source_name: str,
    primary: SourceData,
    style_cache: dict[tuple[int, int], Any],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
    col_span: int,
) -> int:
    if not rows:
        return out_row + 1
    base_col = min(tc.col for tc in plan.cells)
    for i, src_row in enumerate(rows):
        ctx = _build_row_context(
            src_row,
            joined_rows[i],
            join,
            sources,
            active_source_name,
            primary,
            i + 1,
            config_values,
            inputs,
            rows,
        )
        col_offset = i * col_span
        for tc in plan.cells:
            value = _render_cell(tc.template, ctx)
            style = style_cache.get((plan.template_row, tc.col))
            if tc.template.is_single_expression and style is not None:
                value = _apply_numfmt_coercion(value, style[4])
            new_col = tc.col + col_offset
            # First record reuses the original cell column; subsequent records
            # shift by `col_span` per record.
            target = ws.cell(row=out_row, column=new_col, value=value)
            _apply_style(target, style)
            _ = base_col  # currently unused; kept for future left-anchor needs
    return out_row + 1


def _build_row_context(
    src_row: dict[str, Any],
    joined_row: dict[str, Any] | None,
    join: JoinDirective | None,
    sources: dict[str, SourceData],
    active_source_name: str,
    primary: SourceData,
    row_index: int,
    config_values: dict[str, Any],
    inputs: dict[str, Any],
    active_row_set: list[dict[str, Any]] | None,
) -> EvalContext:
    joined_rows: dict[str, dict[str, Any]] = {}
    joined_columns: dict[str, set[str]] = {}
    if join is not None and joined_row is not None:
        joined_rows[join.joined_source] = joined_row
        joined_columns[join.joined_source] = set(sources[join.joined_source].headers)
    return EvalContext(
        active_row=src_row,
        active_source_name=active_source_name,
        active_source_columns=set(primary.headers) if primary.headers else None,
        joined_rows=joined_rows,
        joined_columns=joined_columns,
        inputs=inputs,
        config_values=config_values,
        active_row_set=active_row_set,
        named_sources=_build_named_sources_view(sources),
        row_index=row_index,
    )


def _build_named_sources_view(sources: dict[str, SourceData]) -> dict[str, dict[str, Any]]:
    return {name: {"headers": sd.headers, "rows": sd.rows} for name, sd in sources.items()}


# ---------------------------------------------------------------------------
# numFmt-driven single-expression coercion (ADR-0003)
# ---------------------------------------------------------------------------


def _apply_numfmt_coercion(value: Any, number_format: str | None) -> Any:
    """ADR-0003: single-expression cells whose template cell has a date /
    number / text format MUST coerce the value to that format. Failures
    raise xl3/cell/numfmt-coercion.
    """
    if value is None or number_format is None:
        return value
    nf = number_format
    if nf == "General":
        return value
    # Text format
    if nf == "@":
        if isinstance(value, str):
            return value
        return canonical_string(value)
    nf_lower = nf.lower()
    has_date_token = any(t in nf_lower for t in ("yyyy", "yy", "mm", "dd", "hh", "ss"))
    if has_date_token and not _is_pure_number_format(nf):
        return _coerce_to_date_for_numfmt(value, nf)
    if any(c in nf for c in "0#"):
        return _coerce_to_number_for_numfmt(value, nf)
    return value


def _is_pure_number_format(nf: str) -> bool:
    """Heuristic: format like `0`, `#,##0`, `0.00` is pure numeric, no date."""
    nf_lower = nf.lower()
    if any(t in nf_lower for t in ("yyyy", "yy", "mm", "dd", "hh", "ss")):
        return False
    return any(c in nf for c in "0#")


def _coerce_to_number_for_numfmt(value: Any, nf: str) -> Any:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        from .value_model import parse_number_strict

        n = parse_number_strict(value)
        if n is None:
            raise xtl_error(
                "xl3/cell/numfmt-coercion",
                f'Value cannot be coerced to a number for cell format "{nf}": {value}',
            )
        return n
    raise xtl_error(
        "xl3/cell/numfmt-coercion",
        f'Value cannot be coerced to a number for cell format "{nf}": {canonical_string(value)}',
    )


def _coerce_to_date_for_numfmt(value: Any, nf: str) -> Any:
    from datetime import date, datetime

    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        s = value.strip()
        try:
            return datetime.fromisoformat(s)
        except ValueError as exc:
            raise xtl_error(
                "xl3/cell/numfmt-coercion",
                f'Value cannot be coerced to a date for cell format "{nf}": {value}',
            ) from exc
    raise xtl_error(
        "xl3/cell/numfmt-coercion",
        f'Value cannot be coerced to a date for cell format "{nf}": {canonical_string(value)}',
    )


def _render_cell(tpl: CellTemplate, ctx: EvalContext) -> Any:
    if tpl.is_pure_text:
        seg = tpl.segments[0]
        assert isinstance(seg, TextSegment)
        return seg.text if seg.text != "" else None

    if tpl.is_single_expression:
        seg = tpl.segments[0]
        assert isinstance(seg, ExprSegment)
        return evaluate(seg.expr, ctx)

    out: list[str] = []
    for seg in tpl.segments:
        if isinstance(seg, TextSegment):
            out.append(seg.text)
        elif isinstance(seg, ExprSegment):
            out.append(canonical_string(evaluate(seg.expr, ctx)))
        elif isinstance(seg, DirectiveSegment):
            # Directive cells should have been classified as directive rows
            # in parser.py and thus stripped before render. Reaching here
            # means a mixed cell — emit nothing for the directive segment.
            out.append("")
    return "".join(out)


def _evaluate_filename(
    parsed: ParsedTemplate,
    sources: dict[str, SourceData],
    config_values: dict[str, Any],
    inputs: dict[str, Any],
) -> str:
    pattern = parsed.meta.output_file_pattern or "output.xlsx"
    tpl = _parse_pattern(pattern)
    if tpl.is_pure_text:
        seg = tpl.segments[0]
        return seg.text if isinstance(seg, TextSegment) else pattern
    default_source = sources.get("default")
    row = default_source.rows[0] if default_source and default_source.rows else {}
    headers = set(default_source.headers) if default_source and default_source.headers else None
    ctx = EvalContext(
        active_row=row,
        inputs=inputs,
        config_values=config_values,
        active_source_columns=headers,
    )
    out: list[str] = []
    for seg in tpl.segments:
        if isinstance(seg, TextSegment):
            out.append(seg.text)
        elif isinstance(seg, ExprSegment):
            out.append(canonical_string(evaluate(seg.expr, ctx)))
    return "".join(out)


def _parse_pattern(pattern: str) -> CellTemplate:
    """Parse a filename pattern. Bare identifiers like `{{ Customer }}` are
    treated as group-key references per language.md §"Group Keys".
    """
    from .expression import (
        DirectiveSegment as _DS,
        ExprSegment as _ES,
        TextSegment as _TS,
    )

    # Reuse the `{{ ... }}` splitter but parse each body with the relaxed
    # filename grammar.
    import re as _re

    segs: list[Any] = []
    i = 0
    for m in _re.finditer(r"\{\{(.*?)\}\}", pattern, flags=_re.DOTALL):
        if m.start() > i:
            segs.append(_TS(pattern[i : m.start()]))
        body = m.group(1)
        if body.lstrip().startswith("@"):
            segs.append(_DS(body=body))
        else:
            from .expression import parse_filename_or_sheet_expression as _pf

            segs.append(_ES(expr=_pf(body), body=body))
        i = m.end()
    if i < len(pattern):
        segs.append(_TS(pattern[i:]))
    if not segs:
        segs.append(_TS(""))
    return CellTemplate(segments=segs)
