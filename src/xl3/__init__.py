"""xl3 — XTL (Excel Template Language) 0.1, Python reference implementation.

This module mirrors the public API shape of the TypeScript reference
(`xl3` on npm). Function names use Python conventions (snake_case), but
**error messages and error codes are byte-identical to the TS impl** —
they are part of the conformance contract (ADR-0015).
"""

from __future__ import annotations

from .errors import XtlError, XtlErrorCode, is_xtl_error, xtl_error
from .types import (
    ConvertOptions,
    InputSpec,
    InputType,
    OutputFile,
    PreviewFile,
    PreviewResult,
    PreviewSheet,
    PreviewSource,
    SourceSpec,
)

__version__ = "0.1.0a0"

__all__ = [
    # public API surface
    "convert",
    "preview",
    "read_template_inputs",
    "xtl_error",
    "is_xtl_error",
    # types
    "ConvertOptions",
    "InputSpec",
    "InputType",
    "OutputFile",
    "PreviewFile",
    "PreviewResult",
    "PreviewSheet",
    "PreviewSource",
    "SourceSpec",
    "XtlError",
    "XtlErrorCode",
    "__version__",
]


def convert(
    template: bytes,
    source: bytes,
    options: ConvertOptions | None = None,
) -> list[OutputFile]:
    """Run a full conversion: template + source → list of output files.

    Mirrors the TS `convert(templateBuffer, sourceBuffer, options)` entry.
    """
    from .pipeline import run_convert

    return run_convert(template, source, options)


def preview(
    template: bytes,
    source: bytes,
    options: ConvertOptions | None = None,
) -> PreviewResult:
    """Inspect what `convert()` will produce without rendering full files."""
    from .pipeline import run_preview

    return run_preview(template, source, options)


def read_template_inputs(template: bytes) -> list[InputSpec]:
    """Inspect a template's runtime input declarations (ADR-0010)."""
    from .pipeline import run_read_inputs

    return run_read_inputs(template)
