"""Tests for ``kaos_office.docx.numbering.formatters``.

Locks the wrap-around / boundary behavior of letter and roman
converters that the kelvin-office implementation got subtly right —
these tests are the regression net against future "simplifications"
that would re-introduce off-by-one bugs.
"""

from __future__ import annotations

import pytest

from kaos_office.docx.numbering import (
    BULLET_CHAR,
    format_decimal,
    format_decimal_zero,
    format_lower_letter,
    format_lower_roman,
    format_number,
    format_ordinal,
    format_upper_letter,
    format_upper_roman,
    is_ordered_format,
)


class TestDecimal:
    def test_one(self) -> None:
        assert format_decimal(1) == "1"

    def test_eleven(self) -> None:
        assert format_decimal(11) == "11"

    def test_zero(self) -> None:
        assert format_decimal(0) == "0"


class TestDecimalZero:
    def test_one(self) -> None:
        assert format_decimal_zero(1) == "01"

    def test_nine(self) -> None:
        assert format_decimal_zero(9) == "09"

    def test_ten(self) -> None:
        assert format_decimal_zero(10) == "10"

    def test_one_hundred(self) -> None:
        assert format_decimal_zero(100) == "100"


class TestLowerLetter:
    """The 1-based Excel-column wraparound is load-bearing — every
    boundary explicitly asserted."""

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (1, "a"),
            (2, "b"),
            (25, "y"),
            (26, "z"),
            (27, "aa"),
            (28, "ab"),
            (51, "ay"),
            (52, "az"),
            (53, "ba"),
            (701, "zy"),
            (702, "zz"),
            (703, "aaa"),
            (18278, "zzz"),
            (18279, "aaaa"),
        ],
    )
    def test_known_values(self, value: int, expected: str) -> None:
        assert format_lower_letter(value) == expected

    def test_zero_returns_empty(self) -> None:
        assert format_lower_letter(0) == ""

    def test_negative_returns_empty(self) -> None:
        assert format_lower_letter(-1) == ""


class TestUpperLetter:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (53, "BA"), (702, "ZZ"), (703, "AAA")],
    )
    def test_known_values(self, value: int, expected: str) -> None:
        assert format_upper_letter(value) == expected


class TestLowerRoman:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (1, "i"),
            (2, "ii"),
            (3, "iii"),
            (4, "iv"),
            (5, "v"),
            (9, "ix"),
            (10, "x"),
            (40, "xl"),
            (49, "xlix"),
            (50, "l"),
            (90, "xc"),
            (99, "xcix"),
            (100, "c"),
            (400, "cd"),
            (500, "d"),
            (900, "cm"),
            (1000, "m"),
            (1994, "mcmxciv"),
            (3999, "mmmcmxcix"),
        ],
    )
    def test_known_values(self, value: int, expected: str) -> None:
        assert format_lower_roman(value) == expected

    def test_zero_returns_empty(self) -> None:
        assert format_lower_roman(0) == ""

    def test_negative_returns_empty(self) -> None:
        assert format_lower_roman(-5) == ""


class TestUpperRoman:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "I"), (4, "IV"), (9, "IX"), (49, "XLIX"), (1994, "MCMXCIV")],
    )
    def test_known_values(self, value: int, expected: str) -> None:
        assert format_upper_roman(value) == expected


class TestOrdinal:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (1, "1st"),
            (2, "2nd"),
            (3, "3rd"),
            (4, "4th"),
            (11, "11th"),
            (12, "12th"),
            (13, "13th"),
            (21, "21st"),
            (22, "22nd"),
            (23, "23rd"),
            (101, "101st"),
            (111, "111th"),
            (113, "113th"),
            (121, "121st"),
        ],
    )
    def test_english_ordinal(self, value: int, expected: str) -> None:
        assert format_ordinal(value) == expected


class TestFormatNumberDispatch:
    def test_known_format(self) -> None:
        assert format_number(11, "decimal") == "11"
        assert format_number(11, "lowerLetter") == "k"
        assert format_number(11, "upperRoman") == "XI"

    def test_bullet_returns_glyph(self) -> None:
        assert format_number(99, "bullet") == BULLET_CHAR

    def test_none_returns_empty(self) -> None:
        assert format_number(99, "none") == ""

    def test_unknown_falls_back_to_decimal(self) -> None:
        # Unknown formats log a warning and fall back to decimal so the
        # attorney still sees *something* citable.
        assert format_number(42, "unknown_format_xyz") == "42"


class TestIsOrderedFormat:
    def test_decimal_ordered(self) -> None:
        assert is_ordered_format("decimal") is True

    def test_lower_letter_ordered(self) -> None:
        assert is_ordered_format("lowerLetter") is True

    def test_bullet_not_ordered(self) -> None:
        assert is_ordered_format("bullet") is False

    def test_none_not_ordered(self) -> None:
        assert is_ordered_format("none") is False

    def test_empty_not_ordered(self) -> None:
        assert is_ordered_format("") is False
