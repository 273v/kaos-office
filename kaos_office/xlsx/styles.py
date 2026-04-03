"""Style table parser for XLSX — date format detection.

Excel stores dates as floating-point numbers. The only way to know a number
is a date is to check its numFmtId in styles.xml. Built-in IDs 14-22 and
45-47 are date/time formats. Custom formats (164+) need regex detection.

Also handles Excel serial number → Python date/datetime conversion.
"""

from __future__ import annotations

import datetime
import re
from typing import TYPE_CHECKING

from kaos_office.ooxml.namespace import SML_CELL_XFS, SML_NUM_FMT, SML_NUM_FMTS, SML_XF

if TYPE_CHECKING:
    from lxml.etree import _Element  # ty: ignore[unresolved-import]

# Built-in date/time numFmtIds
_DATE_FMT_IDS = frozenset({14, 15, 16, 17, 22})
_TIME_FMT_IDS = frozenset({18, 19, 20, 21, 45, 46, 47})
_ALL_DATE_FMT_IDS = _DATE_FMT_IDS | _TIME_FMT_IDS

# Heuristic for custom date format codes: contains y, d, h, s or
# m not preceded by h (month vs minute disambiguation)
_DATE_CODE_RE = re.compile(r"[yYdDhHsS]")
_STRIP_LITERALS_RE = re.compile(r'"[^"]*"|\\.')
_STRIP_BRACKETS_RE = re.compile(r"\[[^\]]*\]")

# Excel epoch: 1899-12-30 (serial 1 = Jan 1, 1900)
_EPOCH_1900 = datetime.datetime(1899, 12, 30, tzinfo=None)
_EPOCH_1904 = datetime.datetime(1904, 1, 1, tzinfo=None)


class StyleTable:
    """Parsed style info for date detection from xl/styles.xml."""

    def __init__(self, xml: _Element | None) -> None:
        self._xf_numfmt: list[int] = []
        self._custom_date_ids: set[int] = set()

        if xml is None:
            return

        # Parse custom numFmts (id >= 164)
        numfmts_el = xml.find(SML_NUM_FMTS)
        if numfmts_el is not None:
            for fmt in numfmts_el.iterchildren(SML_NUM_FMT):
                fid = int(fmt.get("numFmtId", "0"))
                code = fmt.get("formatCode", "")
                if _is_date_format_code(code):
                    self._custom_date_ids.add(fid)

        # Parse cellXfs → list of numFmtIds by style index
        xfs_el = xml.find(SML_CELL_XFS)
        if xfs_el is not None:
            for xf in xfs_el.iterchildren(SML_XF):
                self._xf_numfmt.append(int(xf.get("numFmtId", "0")))

    def is_date(self, style_index: int) -> bool:
        """Check if a style index represents a date/time format."""
        if style_index < 0 or style_index >= len(self._xf_numfmt):
            return False
        nf = self._xf_numfmt[style_index]
        return nf in _ALL_DATE_FMT_IDS or nf in self._custom_date_ids

    def is_time_only(self, style_index: int) -> bool:
        """Check if this is a time-only format (h:mm, h:mm:ss, etc.)."""
        if style_index < 0 or style_index >= len(self._xf_numfmt):
            return False
        return self._xf_numfmt[style_index] in _TIME_FMT_IDS


def _is_date_format_code(code: str) -> bool:
    """Heuristic: does this custom format code represent a date/time?"""
    # Strip quoted literals and escaped chars
    stripped = _STRIP_LITERALS_RE.sub("", code)
    # Strip bracket expressions like [Red], [>100]
    stripped = _STRIP_BRACKETS_RE.sub("", stripped)
    return bool(_DATE_CODE_RE.search(stripped))


def serial_to_date(
    serial: float, *, date1904: bool = False
) -> datetime.date | datetime.datetime | datetime.time:
    """Convert Excel serial number to Python date/datetime/time.

    - Integer serial (no fractional) → date
    - Fractional serial → datetime
    - serial < 1.0 → time only
    """
    epoch = _EPOCH_1904 if date1904 else _EPOCH_1900

    if serial < 1.0:
        # Time only
        total_seconds = round(serial * 86400)
        h, rem = divmod(total_seconds, 3600)
        m, s = divmod(rem, 60)
        return datetime.time(h, m, s)

    # Handle the Excel 1900 leap year bug: serial 60 = Feb 29, 1900 (doesn't exist)
    if not date1904 and serial == 60:
        return datetime.date(1900, 2, 28)  # Best approximation
    if not date1904 and serial > 60:
        serial -= 1  # Adjust for the phantom leap day

    days = int(serial)
    frac = serial - days

    dt = epoch + datetime.timedelta(days=days)

    if frac > 0.0001:  # Has time component
        total_seconds = round(frac * 86400)
        return dt + datetime.timedelta(seconds=total_seconds)

    return dt.date()
