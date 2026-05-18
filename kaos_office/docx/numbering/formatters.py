"""Number format converters for DOCX list numbering.

Converts an integer counter value (e.g. ``11``) into the visible numeral
rendered for a given ``numFmt`` (e.g. ``"decimal"`` → ``"11"``,
``"lowerLetter"`` → ``"k"``, ``"upperRoman"`` → ``"XI"``).

The US-legal subset (``decimal``, ``decimalZero``, ``lowerLetter``,
``upperLetter``, ``lowerRoman``, ``upperRoman``, ``ordinal``,
``bullet``, ``none``) is implemented here; other formats fall back to
plain decimal with a structured log warning. International formats can
be added incrementally — see
``kaos-modules/docs/plans/docx-numbering-resolution.md`` Stage 7.
"""

from __future__ import annotations

import string
from collections.abc import Callable
from typing import Final

from kaos_core.logging import get_logger

_log = get_logger(__name__)

# Bullet glyph used by Word for unordered lists (U+2022 BULLET).
BULLET_CHAR: Final[str] = "•"

_ROMAN_VAL_MAP: Final[tuple[tuple[int, str], ...]] = (
    (1000, "m"),
    (900, "cm"),
    (500, "d"),
    (400, "cd"),
    (100, "c"),
    (90, "xc"),
    (50, "l"),
    (40, "xl"),
    (10, "x"),
    (9, "ix"),
    (5, "v"),
    (4, "iv"),
    (1, "i"),
)

_ORDINAL_SUFFIXES: Final[dict[int, str]] = {1: "st", 2: "nd", 3: "rd"}


def format_decimal(value: int) -> str:
    """Format as decimal: 1, 2, 3, 10, 11."""
    return str(value)


def format_decimal_zero(value: int) -> str:
    """Format as two-digit decimal: 01, 02, ..., 09, 10, 11."""
    return f"{value:02d}"


def format_lower_letter(value: int) -> str:
    """Format as lowercase Excel-style letters: a, b, ..., z, aa, ab, ...

    The 1-based wraparound is load-bearing: 26 → ``"z"``, 27 → ``"aa"``,
    52 → ``"az"``, 53 → ``"ba"``, 702 → ``"zz"``, 703 → ``"aaa"``. Do
    not "simplify" the inner ``value -= 1`` — it adjusts each subsequent
    digit back into 1-based space.
    """
    if value <= 0:
        return ""
    result = ""
    value -= 1
    while True:
        result = string.ascii_lowercase[value % 26] + result
        value //= 26
        if value == 0:
            break
        value -= 1
    return result


def format_upper_letter(value: int) -> str:
    """Format as uppercase Excel-style letters. See :func:`format_lower_letter`."""
    return format_lower_letter(value).upper()


def format_lower_roman(value: int) -> str:
    """Format as lowercase Roman numerals: i, ii, iii, iv, v, ..., mcmxciv."""
    if value <= 0:
        return ""
    parts: list[str] = []
    for arabic, roman in _ROMAN_VAL_MAP:
        count, value = divmod(value, arabic)
        if count:
            parts.append(roman * count)
    return "".join(parts)


def format_upper_roman(value: int) -> str:
    """Format as uppercase Roman numerals."""
    return format_lower_roman(value).upper()


def format_ordinal(value: int) -> str:
    """Format as English ordinal: 1st, 2nd, 3rd, 4th, ..., 11th, 12th, 13th, 21st."""
    if value <= 0:
        return str(value)
    # 11-19 always take "th" regardless of last digit.
    last_two = value % 100
    if 11 <= last_two <= 13:
        return f"{value}th"
    suffix = _ORDINAL_SUFFIXES.get(value % 10, "th")
    return f"{value}{suffix}"


def format_bullet(_value: int) -> str:
    """Bullet format: the visible bullet glyph regardless of counter."""
    return BULLET_CHAR


def format_none(_value: int) -> str:
    """No visible numeral."""
    return ""


_FORMATTERS: Final[dict[str, Callable[[int], str]]] = {
    "decimal": format_decimal,
    "decimalZero": format_decimal_zero,
    "lowerLetter": format_lower_letter,
    "upperLetter": format_upper_letter,
    "lowerRoman": format_lower_roman,
    "upperRoman": format_upper_roman,
    "ordinal": format_ordinal,
    "bullet": format_bullet,
    "none": format_none,
}


def format_number(value: int, num_fmt: str) -> str:
    """Format ``value`` as a visible numeral for ``num_fmt``.

    Returns the empty string for ``"none"`` and the bullet glyph for
    ``"bullet"``. Unknown formats log a structured warning and fall
    back to decimal so an attorney never sees a blank citation token
    when Word actually emitted a numeral.
    """
    formatter = _FORMATTERS.get(num_fmt)
    if formatter is not None:
        return formatter(value)
    _log.warning(
        "numbering.unknown_format",
        extra={"num_fmt": num_fmt, "value": value, "fallback": "decimal"},
    )
    return format_decimal(value)


def is_ordered_format(num_fmt: str) -> bool:
    """Return True for any format that renders a counter (i.e. not bullet/none)."""
    return num_fmt not in {"bullet", "none", ""}
