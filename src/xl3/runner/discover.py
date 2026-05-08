"""Fixture discovery + meta.yaml loading.

A fixture is any subdirectory of `--fixture-dir` that has a `meta.yaml`.
Three fixture kinds (mutually exclusive):

- **static**: has `expected.xlsx` (single output) or `expected/` directory
- **error**:  has `expected_error` in meta.yaml; no expected workbook
- **dynamic**: has `expected_dynamic` in meta.yaml; no expected workbook
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

import yaml

FixtureKind = Literal["static", "error", "dynamic"]


@dataclass
class DynamicCell:
    sheet: str
    cell: str
    format: str


@dataclass
class FixtureInputDecl:
    """An `inputs:` row from meta.yaml — host-supplied runtime input."""

    name: str
    value: Any


@dataclass
class Fixture:
    """A loaded conformance fixture."""

    id: str  # e.g. "001-bracket-substitution"
    path: Path
    meta: dict[str, Any]
    kind: FixtureKind

    description: str = ""
    spec_version: str = "0.1"
    tags: list[str] = field(default_factory=list)
    comparison_stage: int = 1

    template_path: Path | None = None
    data_path: Path | None = None

    expected_path: Path | None = None  # single-file expected.xlsx
    expected_dir: Path | None = None  # multi-file expected/

    expected_error: str | None = None
    expected_error_code: str | None = None
    expected_warnings: list[str] = field(default_factory=list)

    expected_dynamic_kind: str | None = None
    dynamic_cells: list[DynamicCell] = field(default_factory=list)

    inputs: list[FixtureInputDecl] = field(default_factory=list)

    skip_reason: str | None = None


class FixtureLoadError(Exception):
    pass


def _required_str(meta: dict[str, Any], key: str, fixture_id: str) -> str:
    if key not in meta:
        raise FixtureLoadError(f"{fixture_id}: meta.yaml missing required field {key!r}")
    v = meta[key]
    if not isinstance(v, str):
        raise FixtureLoadError(f"{fixture_id}: meta.yaml field {key!r} must be a string")
    return v


def load_fixture(fixture_dir: Path) -> Fixture:
    """Read one fixture directory into a Fixture record."""
    fixture_id = fixture_dir.name
    meta_path = fixture_dir / "meta.yaml"
    if not meta_path.exists():
        raise FixtureLoadError(f"{fixture_id}: missing meta.yaml")
    with meta_path.open("r", encoding="utf-8") as f:
        meta = yaml.safe_load(f) or {}
    if not isinstance(meta, dict):
        raise FixtureLoadError(f"{fixture_id}: meta.yaml must be a mapping at the top level")

    template_path = fixture_dir / "template.xlsx"
    data_path = fixture_dir / "data.xlsx"
    expected_xlsx = fixture_dir / "expected.xlsx"
    expected_subdir = fixture_dir / "expected"

    expected_error = meta.get("expected_error")
    expected_error_code = meta.get("expected_error_code")
    expected_dynamic = meta.get("expected_dynamic")

    if expected_error and expected_dynamic:
        raise FixtureLoadError(
            f"{fixture_id}: expected_error and expected_dynamic are mutually exclusive"
        )

    if expected_error:
        kind: FixtureKind = "error"
    elif expected_dynamic:
        kind = "dynamic"
    else:
        kind = "static"

    expected_path: Path | None = None
    expected_dir: Path | None = None
    if kind == "static":
        if expected_xlsx.exists():
            expected_path = expected_xlsx
        elif expected_subdir.exists() and expected_subdir.is_dir():
            expected_dir = expected_subdir
        else:
            raise FixtureLoadError(
                f"{fixture_id}: static fixture missing expected.xlsx and expected/ directory"
            )

    dynamic_cells: list[DynamicCell] = []
    raw_cells = meta.get("dynamic_cells") or []
    if not isinstance(raw_cells, list):
        raise FixtureLoadError(f"{fixture_id}: dynamic_cells must be a list")
    for entry in raw_cells:
        if not isinstance(entry, dict):
            raise FixtureLoadError(f"{fixture_id}: each dynamic_cells entry must be a mapping")
        dynamic_cells.append(
            DynamicCell(
                sheet=str(entry["sheet"]),
                cell=str(entry["cell"]),
                format=str(entry["format"]),
            )
        )

    inputs: list[FixtureInputDecl] = []
    raw_inputs = meta.get("inputs") or []
    if not isinstance(raw_inputs, list):
        raise FixtureLoadError(f"{fixture_id}: inputs must be a list")
    for entry in raw_inputs:
        if not isinstance(entry, dict):
            raise FixtureLoadError(f"{fixture_id}: each inputs entry must be a mapping")
        if "name" not in entry:
            raise FixtureLoadError(f"{fixture_id}: each inputs entry must declare `name`")
        inputs.append(FixtureInputDecl(name=str(entry["name"]), value=entry.get("value")))

    expected_warnings = meta.get("expected_warnings") or []
    if not isinstance(expected_warnings, list):
        raise FixtureLoadError(f"{fixture_id}: expected_warnings must be a list")

    return Fixture(
        id=fixture_id,
        path=fixture_dir,
        meta=meta,
        kind=kind,
        description=str(meta.get("description") or ""),
        spec_version=str(meta.get("spec_version") or "0.1"),
        tags=list(meta.get("tags") or []),
        comparison_stage=int(meta.get("comparison_stage") or 1),
        template_path=template_path if template_path.exists() else None,
        data_path=data_path if data_path.exists() else None,
        expected_path=expected_path,
        expected_dir=expected_dir,
        expected_error=str(expected_error) if expected_error else None,
        expected_error_code=str(expected_error_code) if expected_error_code else None,
        expected_warnings=[str(w) for w in expected_warnings],
        expected_dynamic_kind=str(expected_dynamic) if expected_dynamic else None,
        dynamic_cells=dynamic_cells,
        inputs=inputs,
        skip_reason=str(meta["skip_reason"]) if "skip_reason" in meta else None,
    )


def discover_fixtures(fixture_dir: Path) -> list[Fixture]:
    """Iterate `fixture_dir` and load every fixture (sorted by id)."""
    fixtures: list[Fixture] = []
    for sub in sorted(fixture_dir.iterdir()):
        if not sub.is_dir():
            continue
        if not (sub / "meta.yaml").exists():
            continue
        fixtures.append(load_fixture(sub))
    return fixtures
