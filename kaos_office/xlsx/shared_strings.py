"""Shared string table parser for XLSX.

XLSX deduplicates strings into xl/sharedStrings.xml. Cells with t="s"
reference strings by 0-based index into this table.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from kaos_office.ooxml.namespace import SML_R, SML_SI, SML_T

if TYPE_CHECKING:
    from lxml.etree import _Element  # ty: ignore[unresolved-import]


class SharedStringTable:
    """Parsed shared string table from xl/sharedStrings.xml."""

    def __init__(self, xml: _Element) -> None:
        self._strings: list[str] = []
        for si in xml.iterchildren(SML_SI):
            # Plain string: <si><t>text</t></si>
            t_elem = si.find(SML_T)
            if t_elem is not None:
                self._strings.append(t_elem.text or "")
                continue
            # Rich text: <si><r><t>part1</t></r><r><t>part2</t></r></si>
            parts = []
            for r_elem in si.iterchildren(SML_R):
                t = r_elem.find(SML_T)
                if t is not None and t.text:
                    parts.append(t.text)
            self._strings.append("".join(parts))

    def get(self, index: int) -> str:
        """Get string by 0-based index."""
        return self._strings[index]

    def __len__(self) -> int:
        return len(self._strings)
