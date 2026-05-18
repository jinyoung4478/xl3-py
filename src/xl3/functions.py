"""XTL function implementations (case-insensitive at the call site).

Some functions (aggregates, ROW, XLOOKUP) need access to context the
expression-evaluator doesn't carry — they're handled in the evaluator's
context-aware path. The simple, context-free ones live here.
"""

from __future__ import annotations

import calendar
import math
import re
from datetime import date, datetime, timedelta, timezone
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


_TRIM_WS_RE = re.compile(
    "^[\t\n\v\f\r \u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]+"
    "|[\t\n\v\f\r \u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]+$"
)


def _expect_arity(name: str, args: list[Any], expected: int) -> None:
    if len(args) != expected:
        raise xtl_error(
            "xl3/eval/arity-mismatch",
            f"{name}: expected {expected} arguments, got {len(args)}",
        )


def _parse_finite_float(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        n = float(v)
    elif isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        try:
            n = float(s)
        except ValueError:
            return None
    else:
        return None
    return n if math.isfinite(n) else None


def _parse_integer_arg(v: Any) -> int | None:
    n = _parse_finite_float(v)
    if n is None or not n.is_integer():
        return None
    return int(n)


def _parse_truncated_int_arg(v: Any) -> int | None:
    n = _parse_finite_float(v)
    return None if n is None else int(n)


def _require_date(name: str, v: Any, ordinal: str = "") -> datetime:
    d = _coerce_to_datetime(v)
    if d is None:
        where = f" as the {ordinal} argument" if ordinal else ""
        raise xtl_error(
            "xl3/eval/type-mismatch",
            f"{name} expected a date{where}, got {canonical_string(v)!r}",
        )
    if isinstance(d, datetime):
        if d.tzinfo is not None:
            return d.astimezone(timezone.utc).replace(tzinfo=None)
        return d
    return datetime(d.year, d.month, d.day)


def _add_months(year: int, month: int, months: int) -> tuple[int, int]:
    y_delta, zero_based_month = divmod(month - 1 + months, 12)
    return year + y_delta, zero_based_month + 1


def _last_day_of_month(year: int, month: int, name: str) -> int:
    try:
        return calendar.monthrange(year, month)[1]
    except ValueError as exc:
        raise xtl_error("xl3/eval/type-mismatch", f"{name} produced an invalid date") from exc


def _make_datetime(year: int, month: int, day: int, name: str) -> datetime:
    try:
        return datetime(year, month, day)
    except ValueError as exc:
        raise xtl_error("xl3/eval/type-mismatch", f"{name} produced an invalid date") from exc


def fn_if(args: list[Any]) -> Any:
    """IF(condition, then, else) — ADR-0008 truthiness."""
    _expect_arity("IF", args, 3)
    return args[1] if is_truthy(args[0]) else args[2]


def fn_ifempty(args: list[Any]) -> Any:
    """IFEMPTY(value, fallback) — ADR-0007 empty predicate."""
    _expect_arity("IFEMPTY", args, 2)
    return args[1] if is_empty(args[0]) else args[0]


def fn_round(args: list[Any]) -> float:
    """ROUND(value, places) — Excel half-away-from-zero."""
    _expect_arity("ROUND", args, 2)
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
    _expect_arity("ABS", args, 1)
    return abs(float(args[0]))


def fn_today(args: list[Any]) -> datetime:
    """TODAY() — UTC date at render time, midnight (per ADR-0001/0017)."""
    _expect_arity("TODAY", args, 0)
    now = datetime.now(timezone.utc)
    return datetime(now.year, now.month, now.day)


def fn_year(args: list[Any]) -> int:
    _expect_arity("YEAR", args, 1)
    return _require_date("YEAR()", args[0]).year


def fn_month(args: list[Any]) -> int:
    _expect_arity("MONTH", args, 1)
    return _require_date("MONTH()", args[0]).month


def fn_day(args: list[Any]) -> int:
    _expect_arity("DAY", args, 1)
    return _require_date("DAY()", args[0]).day


def fn_eomonth(args: list[Any]) -> datetime:
    _expect_arity("EOMONTH", args, 2)
    d = _require_date("EOMONTH()", args[0], "1st")
    months = _parse_integer_arg(args[1])
    if months is None:
        raise xtl_error("xl3/eval/type-mismatch", "EOMONTH() expected an integer month offset")
    y, m = _add_months(d.year, d.month, months)
    return _make_datetime(y, m, _last_day_of_month(y, m, "EOMONTH()"), "EOMONTH()")


def fn_edate(args: list[Any]) -> datetime:
    _expect_arity("EDATE", args, 2)
    d = _require_date("EDATE()", args[0], "1st")
    months = _parse_integer_arg(args[1])
    if months is None:
        raise xtl_error("xl3/eval/type-mismatch", "EDATE() expected an integer month offset")
    y, m = _add_months(d.year, d.month, months)
    return _make_datetime(y, m, min(d.day, _last_day_of_month(y, m, "EDATE()")), "EDATE()")


def fn_datedif(args: list[Any]) -> int:
    _expect_arity("DATEDIF", args, 3)
    start = _require_date("DATEDIF()", args[0], "1st")
    end = _require_date("DATEDIF()", args[1], "2nd")
    unit = canonical_string(args[2]).upper()
    if unit not in {"Y", "M", "D"}:
        raise xtl_error("xl3/eval/type-mismatch", 'DATEDIF() unit must be "Y", "M", or "D"')
    sign = 1 if end >= start else -1
    a, b = (start, end) if sign == 1 else (end, start)
    if unit == "D":
        return sign * math.floor((b - a).total_seconds() / 86400)
    years = b.year - a.year
    months = b.month - a.month
    days = b.day - a.day
    if days < 0:
        months -= 1
    if months < 0:
        years -= 1
        months += 12
    return sign * (years if unit == "Y" else years * 12 + months)


def fn_upper(args: list[Any]) -> str:
    _expect_arity("UPPER", args, 1)
    return canonical_string(args[0]).upper()


def fn_lower(args: list[Any]) -> str:
    _expect_arity("LOWER", args, 1)
    return canonical_string(args[0]).lower()


def fn_trim(args: list[Any]) -> str:
    _expect_arity("TRIM", args, 1)
    return _TRIM_WS_RE.sub("", canonical_string(args[0]))


def fn_hyperlink(args: list[Any]) -> dict[str, Any]:
    if len(args) not in (1, 2):
        raise xtl_error(
            "xl3/eval/arity-mismatch",
            f"HYPERLINK expects 1 or 2 arguments, got {len(args)}",
        )
    url = canonical_string(args[0]).strip()
    if url == "":
        raise xtl_error(
            "xl3/eval/type-mismatch",
            "HYPERLINK() url argument must be a non-empty string",
        )
    label = args[1] if len(args) == 2 and args[1] not in (None, "") else url
    text = canonical_string(label)
    return {"__xl3_hyperlink__": url, "text": text}


def fn_date(args: list[Any]) -> datetime:
    _expect_arity("DATE", args, 3)
    year = _parse_truncated_int_arg(args[0])
    month = _parse_truncated_int_arg(args[1])
    day = _parse_truncated_int_arg(args[2])
    if year is None or month is None or day is None or year < 0:
        raise xtl_error("xl3/eval/type-mismatch", "DATE() expected finite numeric components")
    y, m = _add_months(year, month, 0)
    try:
        return _make_datetime(y, m, 1, "DATE()") + timedelta(days=day - 1)
    except (OverflowError, ValueError) as exc:
        raise xtl_error("xl3/eval/type-mismatch", "DATE() produced an invalid date") from exc


def fn_isblank(args: list[Any]) -> bool:
    _expect_arity("ISBLANK", args, 1)
    return is_empty(args[0])


# ---------------------------------------------------------------------------
# TEXT() — language.md §"Text Formatting" minimum table
# ---------------------------------------------------------------------------


def fn_text(args: list[Any]) -> str:
    """TEXT(value, format) — string output for date/number formats.

    XTL 0.1 minimum table: date tokens (YYYY/YY/MM/DD/dd/HH/hh/mm/ss) and
    four number formats (`0`, `#,##0`, `0.00`, `#,##0.00`). Half-away-from-
    zero rounding (same rule as ROUND).
    """
    _expect_arity("TEXT", args, 2)
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
    "YEAR": fn_year,
    "MONTH": fn_month,
    "DAY": fn_day,
    "EOMONTH": fn_eomonth,
    "EDATE": fn_edate,
    "DATEDIF": fn_datedif,
    "UPPER": fn_upper,
    "LOWER": fn_lower,
    "TRIM": fn_trim,
    "HYPERLINK": fn_hyperlink,
    "DATE": fn_date,
    "ISBLANK": fn_isblank,
}


def get_simple_function(name: str):  # type: ignore[no-untyped-def]
    """Return a context-free function by upper-cased name, or None."""
    return _BUILTIN_FUNCTIONS.get(name.upper())
