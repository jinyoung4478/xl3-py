"""Pin the ADR-0007/0008/0009/0017 contract before any rendering code is built.

These tests don't depend on Excel I/O — they exercise the value model in
isolation so a regression here surfaces as a unit test failure, not a
mysterious fixture failure 200 lines downstream.
"""

from __future__ import annotations

from datetime import date, datetime, timezone

import pytest

from xl3.value_model import (
    canonical_date,
    canonical_number,
    canonical_string,
    compare_values,
    is_empty,
    is_truthy,
    parse_number_strict,
    trim_unicode_whitespace,
)


# ---------------------------------------------------------------------------
# ADR-0007 — empty
# ---------------------------------------------------------------------------


class TestIsEmpty:
    @pytest.mark.parametrize(
        "v",
        [None, "", "   ", "\t\n", " ", "　"],  # NBSP, ideographic space
    )
    def test_empty(self, v: object) -> None:
        assert is_empty(v)

    @pytest.mark.parametrize("v", [0, 0.0, False, "0", "false", "x", "​", "﻿"])
    def test_non_empty(self, v: object) -> None:
        # ADR-0007: numbers (incl. 0), Booleans (incl. False), and zero-width
        # characters are NOT empty. "0" and "false" strings are not special.
        assert not is_empty(v)

    def test_nan_inf_empty(self) -> None:
        # Per ADR-0009: non-finite stringifies to "" → empty.
        assert is_empty(float("nan"))
        assert is_empty(float("inf"))
        assert is_empty(float("-inf"))


# ---------------------------------------------------------------------------
# ADR-0008 — truthiness
# ---------------------------------------------------------------------------


class TestIsTruthy:
    @pytest.mark.parametrize("v", [False, 0, 0.0, None, "", "  "])
    def test_falsy(self, v: object) -> None:
        assert not is_truthy(v)

    @pytest.mark.parametrize("v", [True, 1, -1, 1.5, "x", "0", "false", "FALSE"])
    def test_truthy(self, v: object) -> None:
        # No special case for "0" / "false" strings (ADR-0008).
        assert is_truthy(v)


# ---------------------------------------------------------------------------
# ADR-0009 / ADR-0017 — canonical string
# ---------------------------------------------------------------------------


class TestCanonicalString:
    def test_empty_to_empty_string(self) -> None:
        assert canonical_string(None) == ""
        assert canonical_string("   ") == ""
        assert canonical_string(float("nan")) == ""

    def test_booleans(self) -> None:
        assert canonical_string(True) == "TRUE"
        assert canonical_string(False) == "FALSE"

    def test_string_passthrough(self) -> None:
        assert canonical_string("hello") == "hello"
        assert canonical_string("0") == "0"

    def test_date_midnight(self) -> None:
        assert canonical_string(datetime(2026, 5, 8)) == "2026-05-08"
        assert canonical_string(date(2026, 5, 8)) == "2026-05-08"

    def test_datetime_non_midnight(self) -> None:
        assert canonical_string(datetime(2026, 5, 8, 9, 30)) == "2026-05-08T09:30:00"

    def test_datetime_with_tz_normalized_to_utc(self) -> None:
        # ADR-0017: components read in UTC.
        kst = timezone.utcoffset  # noqa: F841 — ensure import not stripped
        from datetime import timedelta

        seoul = timezone(timedelta(hours=9))
        d = datetime(2026, 5, 8, 9, 0, 0, tzinfo=seoul)
        # 09:00 KST == 00:00 UTC → midnight → date-only form
        assert canonical_string(d) == "2026-05-08"


class TestCanonicalNumber:
    """ECMA-262 §6.1.6.1.13 — shortest round-trippable, decimal in [1e-6, 1e21)."""

    @pytest.mark.parametrize(
        "x, expected",
        [
            (0, "0"),
            (-0.0, "0"),
            (1, "1"),
            (-1, "-1"),
            (1.5, "1.5"),
            (100.0, "100"),
            (123456789.0, "123456789"),
            (0.5, "0.5"),
            (0.1, "0.1"),
            (0.0001, "0.0001"),
            (0.00001, "0.00001"),
            (0.000001, "0.000001"),  # at the [1e-6, 1e-7) boundary, decimal
            (0.0000001, "1e-7"),  # 1e-7 → scientific, NO leading-zero pad
            (1e21, "1e+21"),
            (1e20, "100000000000000000000"),
            (-123.456, "-123.456"),
            (0.1 + 0.2, "0.30000000000000004"),  # IEEE 754 floats — not rounded
        ],
    )
    def test_format(self, x: float, expected: str) -> None:
        assert canonical_number(x) == expected

    def test_nan_to_empty_string(self) -> None:
        assert canonical_number(float("nan")) == ""
        assert canonical_number(float("inf")) == ""
        assert canonical_number(float("-inf")) == ""

    def test_bool_rejected(self) -> None:
        # bool is a subclass of int — explicit guard so we don't accidentally
        # produce "1" / "0" instead of "TRUE" / "FALSE".
        with pytest.raises(TypeError):
            canonical_number(True)  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Numeric coercion
# ---------------------------------------------------------------------------


class TestParseNumberStrict:
    @pytest.mark.parametrize(
        "s, expected",
        [
            ("0", 0.0),
            ("1", 1.0),
            ("-1", -1.0),
            ("1.5", 1.5),
            ("  42  ", 42.0),
            ("1e3", 1000.0),
        ],
    )
    def test_parses(self, s: str, expected: float) -> None:
        assert parse_number_strict(s) == expected

    @pytest.mark.parametrize("s", ["", "abc", "1.2.3", "nan", "inf", "1a"])
    def test_rejects(self, s: str) -> None:
        assert parse_number_strict(s) is None

    def test_unicode_minus_rejected(self) -> None:
        # ADR-0009: U+2212 is NOT recognized as a number.
        assert parse_number_strict("−5") is None
        assert parse_number_strict("−5.0") is None

    def test_javascript_infinity_keyword(self) -> None:
        # ECMA Number("Infinity") === Infinity. We accept that, then the
        # value flows as empty per ADR-0009 (since canonical_number(inf) == "").
        assert parse_number_strict("Infinity") == float("inf")


# ---------------------------------------------------------------------------
# Comparison (ADR-0009 + ADR-0017)
# ---------------------------------------------------------------------------


class TestCompareValues:
    def test_both_empty_equal(self) -> None:
        assert compare_values(None, "") == 0
        assert compare_values("   ", None) == 0

    def test_one_empty_is_less(self) -> None:
        assert compare_values(None, 0) == -1
        assert compare_values("anything", None) == 1

    def test_both_numeric_strings(self) -> None:
        # Both parse as finite numbers → numeric comparison.
        assert compare_values("100", "20") == 1  # NOT lexical (would be -1)
        assert compare_values("3.14", "3.14") == 0

    def test_ieee754_quirks_preserved(self) -> None:
        # ADR-0009: 0.1 + 0.2 != 0.3 under numeric comparison.
        assert compare_values(0.1 + 0.2, 0.3) != 0

    def test_booleans(self) -> None:
        assert compare_values(False, True) == -1
        assert compare_values(True, True) == 0

    def test_dates_by_timestamp(self) -> None:
        # A midnight Date and a date-time on the same day compare correctly,
        # not as canonical-string strings.
        a = datetime(2026, 5, 8)
        b = datetime(2026, 5, 8, 9, 30)
        assert compare_values(a, b) == -1

    def test_unicode_minus_falls_through_to_string(self) -> None:
        # ADR-0009 explicit: "−5" (U+2212) and -5 do NOT compare equal.
        assert compare_values("−5", -5) != 0

    def test_string_codepoint_order(self) -> None:
        # No locale collation. ASCII < Hangul.
        assert compare_values("Acme", "가나") == -1


# ---------------------------------------------------------------------------
# Trim
# ---------------------------------------------------------------------------


class TestTrim:
    def test_basic(self) -> None:
        assert trim_unicode_whitespace("  hello  ") == "hello"

    def test_unicode_whitespace(self) -> None:
        # NBSP, ideographic space — White_Space property → stripped
        assert trim_unicode_whitespace(" 　hi ") == "hi"

    def test_zero_width_kept(self) -> None:
        # ADR-0007: zero-width chars are NOT whitespace.
        assert trim_unicode_whitespace("​") == "​"
        assert trim_unicode_whitespace("﻿") == "﻿"


# ---------------------------------------------------------------------------
# Canonical date — UTC discipline
# ---------------------------------------------------------------------------


class TestCanonicalDate:
    def test_naive_treated_as_utc(self) -> None:
        # Naive datetime: no tzinfo handling, components used as-is.
        d = datetime(2026, 5, 8, 0, 0, 0)
        assert canonical_date(d) == "2026-05-08"

    def test_non_midnight(self) -> None:
        d = datetime(2026, 5, 8, 12, 34, 56)
        assert canonical_date(d) == "2026-05-08T12:34:56"

    def test_aware_normalized_to_utc(self) -> None:
        from datetime import timedelta

        kst = timezone(timedelta(hours=9))
        # 2026-05-08 09:00 KST == 2026-05-08 00:00 UTC
        d = datetime(2026, 5, 8, 9, 0, 0, tzinfo=kst)
        assert canonical_date(d) == "2026-05-08"
