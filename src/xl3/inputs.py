"""Resolve runtime inputs (ADR-0010).

The host passes a dict of `{name: value}`; we coerce by declared type and
fill defaults. Missing required inputs are an error.
"""

from __future__ import annotations

from datetime import date, datetime, timezone
from typing import Any

from .errors import xtl_error
from .types import InputSpec
from .value_model import canonical_string, parse_number_strict


def resolve_inputs(
    declarations: list[InputSpec],
    supplied: dict[str, Any] | None,
) -> dict[str, Any]:
    """Return a `{name: coerced_value}` dict.

    Per ADR-0010:
      - text     → host string passed through (canonical-stringified if non-string)
      - number   → "trim, then Number() without producing NaN"
      - date     → parsed by date-format coercion rules
      - select   → host value MUST equal one of declared options after
                   canonical-string normalization
    """
    supplied = supplied or {}
    out: dict[str, Any] = {}
    for spec in declarations:
        if spec.name in supplied:
            raw: Any = supplied[spec.name]
            present = True
        elif spec.default is not None:
            raw = spec.default
            present = True
        else:
            present = False
            raw = None
        if not present:
            raise xtl_error(
                "xl3/inputs/missing-required",
                f'Input "{spec.name}" is required',
            )
        out[spec.name] = _coerce(spec, raw)
    return out


def _coerce(spec: InputSpec, raw: Any) -> Any:
    if spec.type == "text":
        if isinstance(raw, str):
            return raw
        return canonical_string(raw)
    if spec.type == "number":
        n = parse_number_strict(raw)
        if n is None:
            raise xtl_error(
                "xl3/inputs/parse-number",
                f'Input "{spec.name}" cannot be parsed as a number: {canonical_string(raw)!r}',
            )
        return n
    if spec.type == "date":
        d = _coerce_date(raw)
        if d is None:
            raise xtl_error(
                "xl3/inputs/parse-date",
                f'Input "{spec.name}" cannot be parsed as a date: {canonical_string(raw)!r}',
            )
        return d
    if spec.type == "select":
        if not spec.options:
            raise xtl_error(
                "xl3/inputs/missing-options",
                f'Input "{spec.name}" is select-typed but has no options',
            )
        s = canonical_string(raw).strip()
        if s not in spec.options:
            raise xtl_error(
                "xl3/inputs/select-option",
                f'Input "{spec.name}" value {s!r} is not one of {spec.options!r}',
            )
        return s
    raise xtl_error(
        "xl3/inputs/invalid-type",
        f'Input "{spec.name}" has unknown type {spec.type!r}',
    )


def _coerce_date(raw: Any) -> datetime | date | None:
    if isinstance(raw, datetime):
        if raw.tzinfo is not None:
            return raw.astimezone(timezone.utc).replace(tzinfo=None)
        return raw
    if isinstance(raw, date):
        return raw
    if isinstance(raw, (int, float)):
        # Excel serial date base (1900) — out of scope for the bootstrap.
        return None
    if isinstance(raw, str):
        s = raw.strip()
        try:
            return datetime.fromisoformat(s)
        except ValueError:
            return None
    return None
