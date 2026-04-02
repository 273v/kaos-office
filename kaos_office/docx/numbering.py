"""DOCX Numbering Resolver.

Parses numbering.xml and resolves numId + ilvl to list type (ordered/bullet).
Handles the three-level indirection: numId → abstractNumId → level definitions → numFmt.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from kaos_office.ooxml.namespace import W_VAL, W, qn

# Number formats that indicate ordered (numbered) lists
_ORDERED_FORMATS = frozenset(
    {
        "decimal",
        "lowerLetter",
        "upperLetter",
        "lowerRoman",
        "upperRoman",
        "decimalZero",
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


@dataclass
class NumberingResolver:
    """Resolve numId + ilvl to list type (ordered/bullet).

    DOCX numbering uses three-level indirection:
    1. Each list paragraph has numId (instance) + ilvl (indent level)
    2. numId maps to an abstractNumId
    3. abstractNumId defines level formats (lvl → numFmt)

    Some num instances override specific levels of their abstractNum.
    """

    # abstractNumId → {ilvl → numFmt}
    _abstract_nums: dict[str, dict[str, str]] = field(default_factory=dict)
    # numId → abstractNumId
    _num_to_abstract: dict[str, str] = field(default_factory=dict)
    # numId → {ilvl → numFmt} overrides
    _num_overrides: dict[str, dict[str, str]] = field(default_factory=dict)

    @classmethod
    def from_xml(cls, numbering_xml: bytes | None) -> NumberingResolver:
        """Create a NumberingResolver from numbering.xml bytes.

        Args:
            numbering_xml: Raw bytes of numbering.xml, or None if missing.

        Returns:
            NumberingResolver instance.
        """
        if numbering_xml is None:
            return cls()

        from kaos_office.opc.security import parse_xml_safe

        root = parse_xml_safe(numbering_xml)

        abstract_nums: dict[str, dict[str, str]] = {}
        for an in root.iter(qn(W, "abstractNum")):
            an_id = an.get(qn(W, "abstractNumId"))
            if an_id is None:
                continue
            levels: dict[str, str] = {}
            for lvl in an.iter(qn(W, "lvl")):
                ilvl = lvl.get(qn(W, "ilvl")) or "0"
                fmt_el = lvl.find(qn(W, "numFmt"))
                fmt = fmt_el.get(W_VAL) if fmt_el is not None else "decimal"
                levels[ilvl] = fmt
            abstract_nums[an_id] = levels

        num_map: dict[str, str] = {}
        num_overrides: dict[str, dict[str, str]] = {}
        for num in root.iter(qn(W, "num")):
            num_id = num.get(qn(W, "numId"))
            if num_id is None:
                continue

            abstract_ref = num.find(qn(W, "abstractNumId"))
            if abstract_ref is not None:
                num_map[num_id] = abstract_ref.get(W_VAL) or ""

            # Check for level overrides
            overrides: dict[str, str] = {}
            for lvl_override in num.iter(qn(W, "lvlOverride")):
                ilvl = lvl_override.get(qn(W, "ilvl"))
                if ilvl is None:
                    continue
                lvl = lvl_override.find(qn(W, "lvl"))
                if lvl is not None:
                    fmt_el = lvl.find(qn(W, "numFmt"))
                    if fmt_el is not None:
                        overrides[ilvl] = fmt_el.get(W_VAL) or "decimal"
            if overrides:
                num_overrides[num_id] = overrides

        return cls(
            _abstract_nums=abstract_nums,
            _num_to_abstract=num_map,
            _num_overrides=num_overrides,
        )

    def is_ordered(self, num_id: str, ilvl: str = "0") -> bool:
        """Check if a numbering instance at a given level is ordered (not bullet).

        Args:
            num_id: The numId value from the paragraph.
            ilvl: The indent level (default "0").

        Returns:
            True if ordered (decimal, lowerLetter, etc.), False if bullet/none.
        """
        fmt = self.get_format(num_id, ilvl)
        return fmt in _ORDERED_FORMATS

    def get_format(self, num_id: str, ilvl: str = "0") -> str:
        """Get the number format for a numId + ilvl.

        Args:
            num_id: The numId value.
            ilvl: The indent level.

        Returns:
            Number format string (e.g., "decimal", "bullet").
        """
        # Check level overrides first
        overrides = self._num_overrides.get(num_id, {})
        if ilvl in overrides:
            return overrides[ilvl]

        # Resolve numId → abstractNumId → level format
        abstract_id = self._num_to_abstract.get(num_id)
        if abstract_id is None:
            return "bullet"  # Default to bullet if not found

        levels = self._abstract_nums.get(abstract_id, {})
        return levels.get(ilvl, "bullet")

    def has_numbering(self, num_id: str) -> bool:
        """Check if a numId exists in numbering definitions."""
        return num_id in self._num_to_abstract
