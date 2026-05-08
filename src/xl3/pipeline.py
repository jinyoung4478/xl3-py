"""Top-level convert/preview/read_inputs orchestration."""

from __future__ import annotations

from .inputs import resolve_inputs
from .parser import parse_template
from .reader import read_all_sources
from .renderer import _RENDER_LISTS, render
from .types import ConvertOptions, InputSpec, OutputFile, PreviewResult


def run_convert(
    template: bytes,
    source: bytes,
    options: ConvertOptions | None = None,
) -> list[OutputFile]:
    parsed = parse_template(template)
    resolved_inputs = resolve_inputs(
        parsed.inputs,
        options.inputs if options else None,
    )
    sources = read_all_sources(
        source,
        parsed.meta.source_sheet,
        parsed.meta.source_table,
        parsed.sources,
    )
    _RENDER_LISTS.push(parsed.list_sheets)
    try:
        return render(
            parsed,
            sources,
            template,
            config_values=parsed.meta.author_values,
            inputs=resolved_inputs,
        )
    finally:
        _RENDER_LISTS.pop()


def run_preview(
    template: bytes,
    source: bytes,
    options: ConvertOptions | None = None,
) -> PreviewResult:
    raise NotImplementedError("preview pipeline not yet implemented")


def run_read_inputs(template: bytes) -> list[InputSpec]:
    parsed = parse_template(template)
    return parsed.inputs
