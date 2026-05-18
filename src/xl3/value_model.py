"""XTL value model — empty / truthy / canonical-string / compare.

Authoritative ADRs (read together as the value-model contract):
- ADR-0007 empty value definition
- ADR-0008 truthiness
- ADR-0009 comparison and string coercion
- ADR-0017 source value model (extends ADR-0009 with Date branch + UTC discipline)
"""

from __future__ import annotations

import math
from datetime import date, datetime, timezone
from typing import Any

# ---------------------------------------------------------------------------
# Whitespace / empty (ADR-0007)
# ---------------------------------------------------------------------------

# ADR-0007: "ECMAScript String.prototype.trim, equivalent to Unicode White_Space
# property. Zero-width characters (U+200B, U+FEFF) are NOT whitespace."
#
# Python's str.isspace() implements Unicode White_Space exactly. We use it
# directly. (See PORTING_NOTES.md "ECMA trim vs White_Space" for why this is
# safe even though ECMAScript trim historically also stripped U+FEFF.)


def trim_unicode_whitespace(s: str) -> str:
    """Trim Unicode White_Space from both ends. Zero-width chars are kept."""
    # str.strip() uses str.isspace() which matches Unicode White_Space exactly
    # for the classes we care about (ASCII space, NBSP, ideographic space, etc.).
    return s.strip()


def is_empty(v: Any) -> bool:
    """True iff `v` is empty per ADR-0007.

    A value is empty iff it is missing (None) or a string that becomes empty
    after Unicode whitespace trim. Numbers (including 0), Booleans (including
    False), and Dates are NEVER empty.

    NaN and infinities, per ADR-0009/0017, "stringify to '' and flow as empty."
    Excel error sentinels are mapped to empty by the reader (ADR-0017).
    """
    if v is None:
        return True
    if isinstance(v, bool):  # before int — bool is a subclass of int
        return False
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return True
        return False
    if isinstance(v, str):
        return len(trim_unicode_whitespace(v)) == 0
    if isinstance(v, (datetime, date)):
        return False
    # Unknown kind — be conservative: not empty.
    return False


# ---------------------------------------------------------------------------
# Truthiness (ADR-0008)
# ---------------------------------------------------------------------------


def is_truthy(v: Any) -> bool:
    """True iff `v` is truthy per ADR-0008.

    Falsy: Boolean False, the number 0, or empty per ADR-0007.
    Otherwise truthy. There is NO special case for the strings "0" or "false".
    """
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return False  # empty per ADR-0007 → falsy
        return v != 0
    if is_empty(v):
        return False
    return True


# ---------------------------------------------------------------------------
# Canonical number (ECMA-262 Number::toString, shortest round-trippable)
# ---------------------------------------------------------------------------
#
# ADR-0009 specifies the "shortest decimal representation that uniquely
# identifies the value, using `.` as the decimal separator and no scientific
# notation for magnitudes between 1e-4 and 1e21." It cites ECMAScript
# `Number.prototype.toString`.
#
# **Spec text vs ECMA-262 actual behavior diverge.** ECMA-262 §6.1.6.1.13 uses
# a -6 cutoff for scientific notation, not -4. So `(0.00005).toString()` in JS
# is "0.00005" (decimal), not "5e-5". The TS reference impl uses the native
# `Number.prototype.toString`, which has -6 cutoff. To pass the corpus we
# match JS behavior. See PORTING_NOTES.md entry #1.


_INF = float("inf")


def canonical_number(x: float) -> str:
    """Format `x` per ECMA-262 §6.1.6.1.13 Number::toString.

    Shortest round-trippable digit string with JS-style exponent rules:
    decimal form for magnitudes in [1e-6, 1e21), scientific otherwise.
    """
    if isinstance(x, bool):  # bool is a subclass of int — guard up front
        raise TypeError("canonical_number does not accept bool")
    if x != x:  # NaN
        return ""  # ADR-0009: non-finites stringify to ""
    if x == _INF or x == -_INF:
        return ""
    if x == 0:
        return "0"  # also handles -0.0
    if x < 0:
        return "-" + canonical_number(-x)

    # Use Python repr for shortest round-trippable digits.
    s = repr(x)
    if "e" in s:
        mant, exp_str = s.split("e")
        e = int(exp_str)
    else:
        mant = s
        e = 0
    if "." in mant:
        ip, fp = mant.split(".")
    else:
        ip, fp = mant, ""

    all_digits = ip + fp
    lead = 0
    while lead < len(all_digits) - 1 and all_digits[lead] == "0":
        lead += 1
    digits = all_digits[lead:]
    # Decimal exponent n s.t. x = int(digits) * 10^(n - len(digits))
    n = len(ip) + e - lead

    # Drop trailing zeros (ECMA: "k as small as possible")
    while len(digits) > 1 and digits[-1] == "0":
        digits = digits[:-1]
    k = len(digits)

    # ECMA-262 format selection
    if k <= n <= 21:
        return digits + "0" * (n - k)
    if 0 < n <= 21:
        return digits[:n] + "." + digits[n:]
    if -6 < n <= 0:
        return "0." + "0" * (-n) + digits
    sign = "+" if (n - 1) >= 0 else "-"
    if k == 1:
        return digits + "e" + sign + str(abs(n - 1))
    return digits[0] + "." + digits[1:] + "e" + sign + str(abs(n - 1))


# ---------------------------------------------------------------------------
# Canonical date (ADR-0017, UTC)
# ---------------------------------------------------------------------------


def canonical_date(d: datetime | date) -> str:
    """ADR-0017 canonical Date string.

    YYYY-MM-DD when the time component is exactly midnight; otherwise
    YYYY-MM-DDTHH:mm:ss. All fields are read in UTC. Naive datetimes from
    openpyxl are TREATED as UTC (Excel serial dates carry no timezone).
    """
    if isinstance(d, datetime):
        if d.tzinfo is not None:
            d = d.astimezone(timezone.utc).replace(tzinfo=None)
        if d.hour == 0 and d.minute == 0 and d.second == 0 and d.microsecond == 0:
            return f"{d.year:04d}-{d.month:02d}-{d.day:02d}"
        return (
            f"{d.year:04d}-{d.month:02d}-{d.day:02d}"
            f"T{d.hour:02d}:{d.minute:02d}:{d.second:02d}"
        )
    # plain date
    return f"{d.year:04d}-{d.month:02d}-{d.day:02d}"


# ---------------------------------------------------------------------------
# Canonical string (ADR-0009 + ADR-0017)
# ---------------------------------------------------------------------------


def is_hyperlink_marker(v: Any) -> bool:
    return isinstance(v, dict) and "__xl3_hyperlink__" in v


def canonical_string(v: Any) -> str:
    """Canonical string form of `v` per ADR-0009 + ADR-0017."""
    if is_empty(v):
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return canonical_number(float(v))
    if isinstance(v, (datetime, date)):
        return canonical_date(v)
    if isinstance(v, str):
        return v
    if is_hyperlink_marker(v):
        return str(v.get("text") or v.get("__xl3_hyperlink__") or "")
    return str(v)


# ---------------------------------------------------------------------------
# Numeric coercion ("trim, then Number() without producing NaN")
# ---------------------------------------------------------------------------


def parse_number_strict(v: Any) -> float | None:
    """ECMAScript `Number()` semantics over a *trimmed* string.

    Returns the parsed float or None if it would have produced NaN.

    **Unicode minus (U+2212) is NOT recognized.** ADR-0009 nails this down:
    a string with U+2212 falls through to canonical-string comparison.
    """
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return float(v)
    if not isinstance(v, str):
        return None
    s = trim_unicode_whitespace(v)
    if s == "":
        return None
    # ECMAScript Number():
    #   - empty string → 0 (we already returned None — empty is treated as
    #     "not a number" in our caller's path because is_empty short-circuits
    #     comparisons before this point)
    #   - "0x" / "0o" / "0b" prefixes parse as integers
    #   - "Infinity" / "-Infinity" / "+Infinity" recognized
    #   - leading "+" or "-" recognized
    #   - whitespace already trimmed
    # Python float() differs:
    #   - accepts "inf", "nan" (case-insensitive); ECMA does NOT (Number("inf") is NaN).
    #   - does NOT accept hex/oct/bin literals.
    # We handle these divergences explicitly.
    if s.lower() in ("nan", "inf", "+inf", "-inf", "infinity", "+infinity", "-infinity"):
        # JS: Number("Infinity") === Infinity, Number("inf") === NaN.
        if s in ("Infinity", "+Infinity"):
            return _INF
        if s == "-Infinity":
            return -_INF
        return None
    # Hex / octal / binary literals (ECMA accepts; Python float() rejects)
    body = s
    sign = 1.0
    if body.startswith(("+", "-")):
        sign = -1.0 if body[0] == "-" else 1.0
        body = body[1:]
    if len(body) >= 2 and body[0] == "0" and body[1] in ("x", "X", "o", "O", "b", "B"):
        try:
            return sign * float(int(body, 0))
        except ValueError:
            return None
    # Standard decimal / float
    try:
        f = float(s)
    except ValueError:
        return None
    if math.isnan(f):
        return None
    return f


# ---------------------------------------------------------------------------
# Comparison (ADR-0009 + ADR-0017)
# ---------------------------------------------------------------------------


def _compare_canonical_strings(a: str, b: str) -> int:
    """Unicode code-point order; no locale-aware collation."""
    if a == b:
        return 0
    return -1 if a < b else 1


def compare_values(a: Any, b: Any) -> int:
    """Three-way compare per ADR-0009 + ADR-0017.

    Returns -1, 0, or 1. Used by IF()/@filter operators and @sort.
    """
    a_empty = is_empty(a)
    b_empty = is_empty(b)
    # 1. Both empty → equal
    if a_empty and b_empty:
        return 0
    # 2. Exactly one empty → empty < non-empty
    if a_empty:
        return -1
    if b_empty:
        return 1
    # Both Booleans
    if isinstance(a, bool) and isinstance(b, bool):
        return (int(a) > int(b)) - (int(a) < int(b))
    # Both Dates
    if isinstance(a, (datetime, date)) and isinstance(b, (datetime, date)):
        ax = a if isinstance(a, datetime) else datetime(a.year, a.month, a.day)
        bx = b if isinstance(b, datetime) else datetime(b.year, b.month, b.day)
        if ax.tzinfo is not None:
            ax = ax.astimezone(timezone.utc).replace(tzinfo=None)
        if bx.tzinfo is not None:
            bx = bx.astimezone(timezone.utc).replace(tzinfo=None)
        return (ax > bx) - (ax < bx)
    # Both numeric (number, or string parsable as finite number)
    a_num = a if isinstance(a, (int, float)) and not isinstance(a, bool) else parse_number_strict(a)
    b_num = b if isinstance(b, (int, float)) and not isinstance(b, bool) else parse_number_strict(b)
    if a_num is not None and b_num is not None:
        af = float(a_num)
        bf = float(b_num)
        # IEEE 754 numeric equality (per ADR-0009)
        return (af > bf) - (af < bf)
    # Fallback: canonical string code-point order
    return _compare_canonical_strings(canonical_string(a), canonical_string(b))
