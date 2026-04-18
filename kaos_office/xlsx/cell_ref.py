"""Cell reference parsing for XLSX."""

from __future__ import annotations

import re

_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def parse_cell_ref(ref: str) -> tuple[int, int]:
    """Parse Excel cell reference to 0-based (row, col).

    'A1' → (0, 0), 'B2' → (1, 1), 'Z1' → (0, 25), 'AA1' → (0, 26).
    """
    m = _CELL_RE.match(ref.upper())
    if not m:
        msg = f"Invalid cell reference: {ref!r}"
        raise ValueError(msg)
    col = col_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return row, col


def col_to_index(letters: str) -> int:
    """Convert column letters to 0-based index. 'A'→0, 'Z'→25, 'AA'→26."""
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def index_to_col_letters(index: int) -> str:
    """Convert 0-based column index to Excel letters. 0→'A', 25→'Z', 26→'AA'."""
    result = []
    n = index + 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))
