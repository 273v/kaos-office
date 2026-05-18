"""Back-compatible ``NumberingResolver`` shim.

The original ``kaos_office.docx.numbering.NumberingResolver`` exposed
three booleans / format strings on top of a hand-rolled XML parser.
The redesign (Stage 2 of the DOCX numbering resolution plan) moves the
real work into :mod:`kaos_office.docx.numbering.definitions`,
:mod:`...parser`, and :mod:`...state`. This module preserves the old
public surface so existing callers (and the existing test suite)
continue to work unchanged while the reader migrates to the new API.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Final

from kaos_office.docx.numbering.definitions import NumberingDefinitions
from kaos_office.docx.numbering.formatters import is_ordered_format
from kaos_office.docx.numbering.parser import parse_numbering_xml

# Formats that count as "ordered" for the purposes of the legacy
# is_ordered() public method. Mirrors the original public contract
# (bullet/none → False; everything else → True). Kept as a frozenset
# for the constant-time membership check.
_ORDERED_FORMATS: Final[frozenset[str]] = frozenset(
    {
        "decimal",
        "decimalZero",
        "lowerLetter",
        "upperLetter",
        "lowerRoman",
        "upperRoman",
        "ordinal",
        "ordinalText",
        "cardinalText",
        "aiueo",
        "iroha",
        "chineseCounting",
        "chineseCountingThousand",
        "chineseLegalSimplified",
        "ideographDigital",
        "ideographLegalTraditional",
        "ideographTraditional",
        "japaneseCounting",
        "japaneseDigitalTenThousand",
        "japaneseLegal",
        "koreanCounting",
        "koreanDigital",
        "koreanDigital2",
        "koreanLegal",
        "taiwaneseCounting",
        "taiwaneseCountingThousand",
        "taiwaneseDigital",
        "thaiLetters",
        "thaiNumbers",
        "vietnameseCounting",
        "hebrew1",
        "hebrew2",
        "arabicAlpha",
        "arabicAbjad",
        "hindiVowels",
        "hindiConsonants",
        "hindiNumbers",
        "hindiCounting",
    }
)


@dataclass(frozen=True)
class NumberingResolver:
    """Resolve ``numId`` + ``ilvl`` to list type (ordered / bullet).

    Back-compatible API. Internally backed by
    :class:`NumberingDefinitions` and the ``numbering.xml`` parser.
    New code should prefer :class:`NumberingDefinitions` and
    :class:`kaos_office.docx.numbering.state.NumberingState` directly —
    they expose the rendered label, not just a boolean.
    """

    definitions: NumberingDefinitions = field(default_factory=NumberingDefinitions)

    @classmethod
    def from_xml(cls, numbering_xml: bytes | None) -> NumberingResolver:
        """Create a :class:`NumberingResolver` from ``numbering.xml`` bytes."""
        return cls(definitions=parse_numbering_xml(numbering_xml))

    def is_ordered(self, num_id: str, ilvl: str | int = "0") -> bool:
        """True if the numbering instance at ``ilvl`` renders a counter
        (decimal / letter / roman / ordinal / international counter),
        False for bullet / none / unknown.
        """
        fmt = self.get_format(num_id, ilvl)
        return fmt in _ORDERED_FORMATS

    def get_format(self, num_id: str, ilvl: str | int = "0") -> str:
        """Return the ``numFmt`` string for ``(num_id, ilvl)``.

        Returns ``"bullet"`` when the instance is unknown (matches the
        original behavior — readers default to bullet rather than
        crashing on a malformed reference).
        """
        level = _coerce_ilvl(ilvl)
        level_def = self.definitions.get_level_definition(num_id, level)
        if level_def is None:
            return "bullet"
        return level_def.num_format

    def has_numbering(self, num_id: str) -> bool:
        """True if ``num_id`` has a resolvable abstract definition."""
        return self.definitions.has_num_id(num_id)


def _coerce_ilvl(ilvl: str | int) -> int:
    if isinstance(ilvl, int):
        return ilvl
    try:
        return int(ilvl)
    except ValueError:
        return 0


# ``is_ordered_format`` is re-exported from formatters for callers that
# imported it from the legacy module location.
__all__ = ["NumberingResolver", "is_ordered_format"]
