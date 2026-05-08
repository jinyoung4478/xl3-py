"""Public dataclasses (mirroring the TS `types.ts` shape).

Only the surface that hosts touch is here. Internal AST/IR types live next to
their owning modules.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

InputType = Literal["text", "number", "date", "select"]


@dataclass
class InputSpec:
    """Declared runtime input from `__inputs__` (ADR-0010)."""

    name: str
    type: InputType
    default: str | None = None
    label: str | None = None
    description: str | None = None
    options: list[str] | None = None


@dataclass
class SourceSpec:
    """Declared named source from `__sources__` (ADR-0012)."""

    name: str
    sheet: str
    table: str = "1"
    description: str | None = None


@dataclass
class OutputFile:
    """A rendered xlsx file. `data` is the .xlsx bytes."""

    filename: str
    data: bytes


@dataclass
class PreviewSheet:
    name: str
    row_count: int


@dataclass
class PreviewFile:
    filename: str
    sheets: list[PreviewSheet]


@dataclass
class PreviewSource:
    name: str
    sheet: str
    table: str
    row_count: int
    headers: list[str]
    description: str | None = None


@dataclass
class PreviewResult:
    files: list[PreviewFile]
    inputs: list[InputSpec]
    sources: list[PreviewSource]
    warnings: list[str] = field(default_factory=list)


@dataclass
class ConvertOptions:
    """Optional host-supplied data for conversion (ADR-0010)."""

    inputs: dict[str, Any] | None = None
