"""DOCX numbering: parse ``numbering.xml`` and resolve list labels.

Public surface:

* :class:`NumberingResolver` — back-compatible boolean / format-string
  API used by readers that only need ordered-vs-bullet.
* :class:`NumberingDefinitions` — the parsed numbering schema. Resolves
  ``(num_id, ilvl)`` → :class:`LevelDefinition`, honoring
  ``<w:lvlOverride>`` and ``<w:startOverride>``.
* :class:`NumberingState` — running counter machine. Emits the
  rendered visible label (``"11."``, ``"(a)"``, ``"11(a)(i)"``) for
  each numbered paragraph as the reader streams the document.
* :func:`parse_numbering_xml` — turn ``word/numbering.xml`` bytes into
  a :class:`NumberingDefinitions`.
* :func:`format_number` — format a single counter value for a given
  ``numFmt`` (``"decimal"``, ``"lowerLetter"``, ``"upperRoman"``, …).

See ``kaos-modules/docs/plans/docx-numbering-resolution.md`` for the
full design.
"""

from __future__ import annotations

from kaos_office.docx.numbering.definitions import (
    AbstractNum,
    LevelDefinition,
    LevelOverride,
    NumberingDefinitions,
    NumInstance,
)
from kaos_office.docx.numbering.formatters import (
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
from kaos_office.docx.numbering.parser import parse_numbering_xml
from kaos_office.docx.numbering.resolver import NumberingResolver
from kaos_office.docx.numbering.state import NumberingState

__all__ = [
    "BULLET_CHAR",
    "AbstractNum",
    "LevelDefinition",
    "LevelOverride",
    "NumInstance",
    "NumberingDefinitions",
    "NumberingResolver",
    "NumberingState",
    "format_decimal",
    "format_decimal_zero",
    "format_lower_letter",
    "format_lower_roman",
    "format_number",
    "format_ordinal",
    "format_upper_letter",
    "format_upper_roman",
    "is_ordered_format",
    "parse_numbering_xml",
]
