"""Template workbook parser.

Reads a template `.xlsx` and returns a `ParsedTemplate` containing:
- `__config__` metadata (source_sheet, source_table, output_file_pattern, ...)
- author-defined `__config__` values
- `__inputs__` declarations (ADR-0010)
- `__sources__` declarations (ADR-0012) — header parsed; resolution happens at read time
- `__lists__` columns (ADR-0011)
- one `SheetTemplate` per visible/template sheet, with rows already grouped
  into blocks (directive rows + their data row, or static rows).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.cell.rich_text import CellRichText

from .directives import (
    BlockDirectives,
    Directive,
    DirectiveParseError,
    parse_directive,
)
from .errors import xtl_error
from .expression import (
    CellTemplate,
    DirectiveSegment,
    ExprSegment,
    collect_referenced_columns,
    expression_has_per_row_ref,
    parse_cell_template,
)
from .types import InputSpec, SourceSpec
from .value_model import canonical_string, is_empty


# ---------------------------------------------------------------------------
# Parsed model
# ---------------------------------------------------------------------------


@dataclass
class TemplateMeta:
    """The system rows of `__config__`."""

    name: str | None = None
    description: str | None = None
    source_sheet: str | None = None
    source_table: str = "1"
    output_file_pattern: str = "output.xlsx"
    match_pattern: str | None = None
    # Author-defined values: any non-system key.
    author_values: dict[str, Any] = field(default_factory=dict)


_SYSTEM_KEYS = {
    "name",
    "description",
    "source_sheet",
    "source_table",
    "output_file_pattern",
    "match_pattern",
}


@dataclass
class TemplateCell:
    """A parsed template cell at a specific (row, col) coordinate."""

    row: int  # 1-based
    col: int  # 1-based
    template: CellTemplate
    referenced_columns: set[str] = field(default_factory=set)
    has_per_row_ref: bool = False
    raw_text: str = ""

    @property
    def has_data_refs(self) -> bool:
        return self.has_per_row_ref

    @property
    def is_directive_cell(self) -> bool:
        return self.template.is_directive_cell


@dataclass
class StaticRowPlan:
    """A row that is emitted verbatim per render — no expansion."""

    template_row: int
    cells: list[TemplateCell]


@dataclass
class DataRowPlan:
    """A row that is expanded once per filtered/sorted source row."""

    template_row: int
    cells: list[TemplateCell]
    directives: BlockDirectives = field(default_factory=BlockDirectives)


@dataclass
class SheetTemplate:
    """One non-reserved sheet from the template workbook.

    `original_name` is the literal sheet name; group keys (bare identifiers in
    the sheet name like `Sheet_{Customer}`) aren't supported in the bootstrap.
    `directive_only_rows` are rows that contained only directive cells —
    those rows are stripped from output (renderer needs to know which rows to
    blank in the template before re-emitting).
    """

    original_name: str
    plan: list[StaticRowPlan | DataRowPlan] = field(default_factory=list)
    max_col: int = 0
    directive_only_rows: set[int] = field(default_factory=set)


@dataclass
class ParsedTemplate:
    meta: TemplateMeta
    sheets: list[SheetTemplate]
    inputs: list[InputSpec] = field(default_factory=list)
    sources: list[SourceSpec] = field(default_factory=list)
    list_sheets: dict[str, list[str]] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Cell text extraction (handles rich-text concatenation per spec)
# ---------------------------------------------------------------------------


def _cell_effective_text(cell: Cell) -> str:
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, CellRichText):
        return "".join(str(part) for part in v)
    if isinstance(v, str):
        return v
    return str(v)


# ---------------------------------------------------------------------------
# Reserved sheet detection (ADR-0011)
# ---------------------------------------------------------------------------


def is_reserved_sheet(name: str) -> bool:
    return (
        len(name) >= 5
        and name.startswith("__")
        and name.endswith("__")
        and name[2:-2].isalpha()
        and name[2:-2].islower()
    )


# ---------------------------------------------------------------------------
# Parser entry
# ---------------------------------------------------------------------------


def parse_template(template_bytes: bytes) -> ParsedTemplate:
    """Parse a template workbook. Read with `data_only=True` so formula
    cells expose their cached results (per ADR-0017)."""
    wb = load_workbook(BytesIO(template_bytes), data_only=True, rich_text=True)

    meta = TemplateMeta()
    if "__config__" in wb.sheetnames:
        meta = _parse_config_sheet(wb["__config__"])

    list_sheets: dict[str, list[str]] = {}
    if "__lists__" in wb.sheetnames:
        list_sheets = _parse_lists_sheet(wb["__lists__"])

    sources: list[SourceSpec] = []
    if "__sources__" in wb.sheetnames:
        sources = _parse_sources_sheet(wb["__sources__"])

    inputs: list[InputSpec] = []
    if "__inputs__" in wb.sheetnames:
        inputs = _parse_inputs_sheet(wb["__inputs__"])

    # ADR-0011: any sheet matching `^__[a-z]+__$` is engine-reserved. The
    # four known names (__config__/__inputs__/__sources__/__lists__) are
    # processed above. Any OTHER dunder-wrapped sheet is an author error.
    _KNOWN_RESERVED = {"__config__", "__inputs__", "__sources__", "__lists__"}
    sheets: list[SheetTemplate] = []
    for sn in wb.sheetnames:
        if is_reserved_sheet(sn):
            if sn not in _KNOWN_RESERVED:
                raise xtl_error(
                    "xl3/sheet/reserved-name",
                    f'Sheet name "{sn}" is reserved (matches __<name>__ pattern)',
                )
            continue
        sheets.append(_parse_sheet_template(wb[sn]))

    return ParsedTemplate(
        meta=meta,
        sheets=sheets,
        inputs=inputs,
        sources=sources,
        list_sheets=list_sheets,
    )


def _parse_config_sheet(ws: Any) -> TemplateMeta:
    meta = TemplateMeta()
    for row in ws.iter_rows(values_only=False):
        if not row or len(row) < 2:
            continue
        key_cell, val_cell = row[0], row[1]
        key = _cell_effective_text(key_cell).strip()
        if not key:
            continue
        val = val_cell.value
        if key in _SYSTEM_KEYS:
            if key == "name":
                meta.name = str(val) if val is not None else None
            elif key == "description":
                meta.description = str(val) if val is not None else None
            elif key == "source_sheet":
                meta.source_sheet = str(val) if val is not None else None
            elif key == "source_table":
                meta.source_table = str(val) if val is not None else "1"
            elif key == "output_file_pattern":
                meta.output_file_pattern = str(val) if val is not None else "output.xlsx"
            elif key == "match_pattern":
                meta.match_pattern = str(val) if val is not None else None
        else:
            meta.author_values[key] = val
    return meta


def _parse_lists_sheet(ws: Any) -> dict[str, list[str]]:
    """Read `__lists__` per ADR-0011: row 1 = list names, columns below = values.

    Each value is canonicalized + trimmed; empty entries are dropped.
    """
    out: dict[str, list[str]] = {}
    if ws.max_row is None or ws.max_row < 1:
        return out
    header_row = list(ws[1])
    seen: set[str] = set()
    columns: list[tuple[int, str]] = []
    for cell in header_row:
        v = cell.value
        if v is None or v == "":
            continue
        name = str(v).strip()
        if not name:
            continue
        if name in seen:
            raise xtl_error(
                "xl3/sheet/duplicate-list-name",
                f'__lists__ has duplicate list name "{name}"',
            )
        seen.add(name)
        columns.append((cell.column, name))
    for col_idx, name in columns:
        values: list[str] = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            if isinstance(v, CellRichText):
                v = "".join(str(part) for part in v)
            s = canonical_string(v).strip()
            if s == "":
                continue
            values.append(s)
        out[name] = values
    return out


def _parse_sources_sheet(ws: Any) -> list[SourceSpec]:
    """Read `__sources__` per ADR-0012. Returns SourceSpec list."""
    if ws.max_row is None or ws.max_row < 1:
        return []
    header_row = list(ws[1])
    header_map: dict[str, int] = {}
    for cell in header_row:
        v = cell.value
        if v is None:
            continue
        name = str(v).strip().lower()
        if name:
            header_map[name] = cell.column
    if "name" not in header_map or "sheet" not in header_map:
        raise xtl_error(
            "xl3/source/missing-header",
            "__sources__ must declare 'name' and 'sheet' columns",
        )
    name_col = header_map["name"]
    sheet_col = header_map["sheet"]
    table_col = header_map.get("table")
    desc_col = header_map.get("description")

    out: list[SourceSpec] = []
    seen: set[str] = set()
    for r in range(2, ws.max_row + 1):
        nm = ws.cell(row=r, column=name_col).value
        sh = ws.cell(row=r, column=sheet_col).value
        if nm is None and sh is None:
            continue
        if nm is None or sh is None:
            raise xtl_error(
                "xl3/source/missing-required",
                f"__sources__ row {r} missing required name/sheet",
            )
        name = str(nm).strip()
        if not name or name == "default" or name.startswith("__"):
            raise xtl_error(
                "xl3/source/invalid-name",
                f'__sources__ row {r} has invalid name "{name}"',
            )
        if name in seen:
            raise xtl_error(
                "xl3/source/duplicate-name",
                f'__sources__ has duplicate source name "{name}"',
            )
        seen.add(name)
        table = "1"
        if table_col is not None:
            tv = ws.cell(row=r, column=table_col).value
            if tv is not None:
                table = str(tv).strip() or "1"
        desc = None
        if desc_col is not None:
            dv = ws.cell(row=r, column=desc_col).value
            if dv is not None:
                desc = str(dv)
        out.append(SourceSpec(name=name, sheet=str(sh).strip(), table=table, description=desc))
    return out


_INPUT_TYPES = {"text", "number", "date", "select"}


def _parse_inputs_sheet(ws: Any) -> list[InputSpec]:
    """Read `__inputs__` per ADR-0010."""
    if ws.max_row is None or ws.max_row < 1:
        return []
    header_row = list(ws[1])
    header_map: dict[str, int] = {}
    for cell in header_row:
        v = cell.value
        if v is None:
            continue
        name = str(v).strip().lower()
        if name:
            header_map[name] = cell.column
    if "name" not in header_map or "type" not in header_map:
        raise xtl_error(
            "xl3/inputs/missing-header",
            "__inputs__ must declare 'name' and 'type' columns",
        )
    name_col = header_map["name"]
    type_col = header_map["type"]
    default_col = header_map.get("default")
    label_col = header_map.get("label")
    desc_col = header_map.get("description")
    options_col = header_map.get("options")

    out: list[InputSpec] = []
    seen: set[str] = set()
    for r in range(2, ws.max_row + 1):
        nm = ws.cell(row=r, column=name_col).value
        tp = ws.cell(row=r, column=type_col).value
        if nm is None and tp is None:
            continue
        if nm is None:
            continue
        name = str(nm).strip()
        if not name:
            continue
        if name in seen:
            raise xtl_error(
                "xl3/inputs/duplicate-name",
                f'__inputs__ has duplicate input name "{name}"',
            )
        seen.add(name)
        type_str = str(tp).strip().lower() if tp is not None else ""
        if type_str not in _INPUT_TYPES:
            raise xtl_error(
                "xl3/inputs/invalid-type",
                f'__inputs__ row {r}: type must be one of text/number/date/select',
            )
        default = None
        if default_col is not None:
            dv = ws.cell(row=r, column=default_col).value
            if dv is not None and str(dv) != "":
                default = str(dv)
        label = None
        if label_col is not None:
            lv = ws.cell(row=r, column=label_col).value
            if lv is not None:
                label = str(lv)
        description = None
        if desc_col is not None:
            ev = ws.cell(row=r, column=desc_col).value
            if ev is not None:
                description = str(ev)
        options = None
        if options_col is not None:
            ov = ws.cell(row=r, column=options_col).value
            if ov is not None and str(ov) != "":
                options = [opt.strip() for opt in str(ov).split("|") if opt.strip()]
        if type_str == "select" and not options:
            raise xtl_error(
                "xl3/inputs/missing-options",
                f'__inputs__ row {r}: select inputs require options',
            )
        out.append(
            InputSpec(
                name=name,
                type=type_str,  # type: ignore[arg-type]
                default=default,
                label=label,
                description=description,
                options=options,
            )
        )
    return out


def _directive_error_code(body: str) -> str:
    """Pick a stable error code based on which directive failed to parse.

    Specific codes are required for fixtures that assert error_code (082,
    094, etc.). Use a generic bucket otherwise.
    """
    s = body.lstrip().lower()
    if s.startswith("@join"):
        return "xl3/join/bad-on-clause"
    if s.startswith("@source"):
        return "xl3/source/undeclared"
    return "xl3/cell/numfmt-coercion"


def _parse_sheet_template(ws: Any) -> SheetTemplate:
    """Walk a template sheet row-by-row, classifying each row as static,
    directive, or data, and group consecutive directive rows + their data row
    into a single DataRowPlan.
    """
    st = SheetTemplate(original_name=ws.title)
    rows_cells: dict[int, list[TemplateCell]] = {}
    if ws.max_row is None:
        return st
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            text = _cell_effective_text(cell)
            if text == "":
                continue
            tpl = parse_cell_template(text)
            refs: set[str] = set()
            has_per_row = False
            for seg in tpl.segments:
                if isinstance(seg, ExprSegment):
                    refs |= collect_referenced_columns(seg.expr)
                    if expression_has_per_row_ref(seg.expr):
                        has_per_row = True
            tc = TemplateCell(
                row=cell.row,
                col=cell.column,
                template=tpl,
                referenced_columns=refs,
                has_per_row_ref=has_per_row,
                raw_text=text,
            )
            rows_cells.setdefault(cell.row, []).append(tc)
            st.max_col = max(st.max_col, cell.column)

    pending: BlockDirectives = BlockDirectives()
    for r in sorted(rows_cells.keys()):
        cells = rows_cells[r]
        is_directive_row = all(c.is_directive_cell for c in cells)
        has_data = any(c.has_data_refs for c in cells)
        if is_directive_row:
            st.directive_only_rows.add(r)
            for c in cells:
                seg = c.template.segments[0]
                assert isinstance(seg, DirectiveSegment)
                try:
                    pending.add(parse_directive(seg.body))
                except DirectiveParseError as e:
                    code = _directive_error_code(seg.body)
                    raise xtl_error(code, f"{e}") from e
            continue
        if has_data:
            st.plan.append(
                DataRowPlan(template_row=r, cells=cells, directives=pending)
            )
            pending = BlockDirectives()
        else:
            st.plan.append(StaticRowPlan(template_row=r, cells=cells))
    return st
