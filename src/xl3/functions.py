"""XTL function implementations (case-insensitive at the call site).

Some functions (aggregates, ROW, XLOOKUP) need access to context the
expression-evaluator doesn't carry — they're handled in the evaluator's
context-aware path. The simple, context-free ones live here.
"""

from __future__ import annotations

import math
from datetime import date, datetime, timezone
from typing import Any

from .errors import xtl_error
from .value_model import (
    canonical_date,
    canonical_number,
    canonical_string,
    compare_values,
    is_empty,
    is_truthy,
    parse_number_strict,
)


def fn_if(args: list[Any]) -> Any:
    """IF(condition, then, else) — ADR-0008 truthiness."""
    if len(args) != 3:
        raise xtl_error(
            "xl3/cell/numfmt-coercion",  # placeholder — TODO: dedicated arity error
            f"IF expects 3 arguments, got {len(args)}",
        )
    return args[1] if is_truthy(args[0]) else args[2]


def fn_ifempty(args: list[Any]) -> Any:
    """IFEMPTY(value, fallback) — ADR-0007 empty predicate."""
    if len(args) != 2:
        raise xtl_error(
            "xl3/cell/numfmt-coercion",
            f"IFEMPTY expects 2 arguments, got {len(args)}",
        )
    return args[1] if is_empty(args[0]) else args[0]


def fn_round(args: list[Any]) -> float:
    """ROUND(value, places) — Excel half-away-from-zero."""
    if len(args) != 2:
        raise xtl_error(
            "xl3/cell/numfmt-coercion",
            f"ROUND expects 2 arguments, got {len(args)}",
        )
    v = float(args[0]) if not isinstance(args[0], bool) else float(int(args[0]))
    places = int(args[1])
    factor = 10**places
    scaled = v * factor
    # half-away-from-zero
    if scaled >= 0:
        rounded = math.floor(scaled + 0.5)
    else:
        rounded = -math.floor(-scaled + 0.5)
    return rounded / factor


def fn_abs(args: list[Any]) -> float:
    if len(args) != 1:
        raise xtl_error("xl3/cell/numfmt-coercion", f"ABS expects 1 argument, got {len(args)}")
    return abs(float(args[0]))


def fn_today(args: list[Any]) -> datetime:
    """TODAY() — UTC date at render time, midnight (per ADR-0001/0017)."""
    if len(args) != 0:
        raise xtl_error("xl3/cell/numfmt-coercion", "TODAY() takes no arguments")
    now = datetime.now(timezone.utc)
    return datetime(now.year, now.month, now.day)


# ---------------------------------------------------------------------------
# TEXT() — language.md §"Text Formatting" minimum table
# ---------------------------------------------------------------------------


def fn_text(args: list[Any]) -> str:
    """TEXT(value, format) — string output for date/number formats.

    XTL 0.1 minimum table: date tokens (YYYY/YY/MM/DD/dd/HH/hh/mm/ss) and
    four number formats (`0`, `#,##0`, `0.00`, `#,##0.00`). Half-away-from-
    zero rounding (same rule as ROUND).
    """
    if len(args) != 2:
        raise xtl_error("xl3/cell/numfmt-coercion", f"TEXT expects 2 arguments, got {len(args)}")
    value, fmt = args[0], str(args[1] if args[1] is not None else "")
    # Number format paths
    if fmt in ("0", "#,##0", "0.00", "#,##0.00"):
        return _format_number(value, fmt)
    # Date format path: ALL date tokens are uppercase letters; if the format
    # contains a date token character we treat as date-format.
    if any(tok in fmt for tok in ("YYYY", "YY", "MM", "DD", "dd", "HH", "hh", "mm", "ss")):
        return _format_date_text(value, fmt)
    # Fallthrough: implementation-defined extension formats — the spec says
    # implementations MAY accept additional formats but their output is
    # implementation-defined. Return the value as a canonical string for
    # now; portable templates won't depend on this.
    return canonical_string(value)


def _format_number(value: Any, fmt: str) -> str:
    """Half-away-from-zero rounding + JS-compatible formatting.

    Handles the four XTL 0.1 number formats:
      "0"         → integer rounding, no separators
      "#,##0"     → integer rounding, comma thousand separators
      "0.00"      → two decimals, no separators
      "#,##0.00"  → two decimals, comma thousand separators
    """
    if isinstance(value, bool):
        n: float = 1.0 if value else 0.0
    elif isinstance(value, (int, float)):
        n = float(value)
    else:
        parsed = parse_number_strict(value)
        if parsed is None:
            raise xtl_error(
                "xl3/cell/numfmt-coercion",
                f"TEXT cannot coerce {canonical_string(value)!r} to a number",
            )
        n = parsed

    decimals = 2 if "." in fmt else 0
    grouped = "," in fmt
    rounded = _round_half_away_from_zero(n, decimals)
    sign = "-" if rounded < 0 else ""
    abs_val = abs(rounded)
    if decimals == 0:
        int_str = f"{int(round(abs_val)):d}"
        if grouped:
            int_str = _add_thousands(int_str)
        return sign + int_str
    s = f"{abs_val:.{decimals}f}"
    int_part, frac_part = s.split(".")
    if grouped:
        int_part = _add_thousands(int_part)
    return sign + int_part + "." + frac_part


def _round_half_away_from_zero(n: float, decimals: int) -> float:
    factor = 10**decimals
    scaled = n * factor
    if scaled >= 0:
        rounded = math.floor(scaled + 0.5)
    else:
        rounded = -math.floor(-scaled + 0.5)
    return rounded / factor


def _add_thousands(digits: str) -> str:
    if len(digits) <= 3:
        return digits
    out_parts: list[str] = []
    rem = digits
    while len(rem) > 3:
        out_parts.append(rem[-3:])
        rem = rem[:-3]
    out_parts.append(rem)
    return ",".join(reversed(out_parts))


def _format_date_text(value: Any, fmt: str) -> str:
    """Render a date value through XTL date tokens. Reads components in UTC
    per ADR-0017 (naive datetimes treated as UTC-anchored)."""
    d = _coerce_to_datetime(value)
    if d is None:
        raise xtl_error(
            "xl3/cell/numfmt-coercion",
            f"TEXT cannot coerce {canonical_string(value)!r} to a date",
        )
    if isinstance(d, datetime) and d.tzinfo is not None:
        d = d.astimezone(timezone.utc).replace(tzinfo=None)
    elif isinstance(d, date) and not isinstance(d, datetime):
        d = datetime(d.year, d.month, d.day)

    out: list[str] = []
    i = 0
    while i < len(fmt):
        if fmt.startswith("YYYY", i):
            out.append(f"{d.year:04d}")
            i += 4
        elif fmt.startswith("YY", i):
            out.append(f"{d.year % 100:02d}")
            i += 2
        elif fmt.startswith("MM", i):
            out.append(f"{d.month:02d}")
            i += 2
        elif fmt.startswith("DD", i) or fmt.startswith("dd", i):
            out.append(f"{d.day:02d}")
            i += 2
        elif fmt.startswith("HH", i):
            out.append(f"{d.hour:02d}")
            i += 2
        elif fmt.startswith("hh", i):
            h12 = d.hour % 12 or 12
            out.append(f"{h12:02d}")
            i += 2
        elif fmt.startswith("mm", i):
            out.append(f"{d.minute:02d}")
            i += 2
        elif fmt.startswith("ss", i):
            out.append(f"{d.second:02d}")
            i += 2
        else:
            out.append(fmt[i])
            i += 1
    return "".join(out)


def _coerce_to_datetime(v: Any) -> datetime | date | None:
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip()
        # ISO date forms accepted by Python: "YYYY-MM-DD" and "YYYY-MM-DDTHH:MM:SS".
        try:
            return datetime.fromisoformat(s)
        except ValueError:
            return None
    return None


# Stubs for functions that need richer context — wired up later.
_BUILTIN_FUNCTIONS = {
    "IF": fn_if,
    "IFEMPTY": fn_ifempty,
    "ROUND": fn_round,
    "ABS": fn_abs,
    "TEXT": fn_text,
    "TODAY": fn_today,
}


def get_simple_function(name: str):  # type: ignore[no-untyped-def]
    """Return a context-free function by upper-cased name, or None."""
    return _BUILTIN_FUNCTIONS.get(name.upper())
