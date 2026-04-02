"""DOCX Style Resolver.

Parses styles.xml and resolves paragraph styles to heading levels.
Walks inheritance chains (basedOn) with cycle detection.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from lxml import etree

from kaos_office.ooxml.namespace import W_BASED_ON, W_NAME, W_OUTLINE_LVL, W_PPR, W_VAL, W, qn

_HEADING_NAME_RE = re.compile(r"^heading\s+(\d+)$", re.IGNORECASE)
_TOC_HEADING_RE = re.compile(r"^toc\s*heading$", re.IGNORECASE)


@dataclass
class StyleResolver:
    """Resolve paragraph styles to heading levels and detect style properties.

    Parses styles.xml into a dict of style elements, then resolves heading
    levels via three mechanisms (checked in order):
    1. outlineLvl in the style's pPr
    2. Style name matching "Heading N"
    3. Walking the basedOn inheritance chain
    """

    _styles: dict[str, etree._Element] = field(default_factory=dict)
    _heading_cache: dict[str, int | None] = field(default_factory=dict)
    _seen: set[str] = field(default_factory=set)  # Cycle detection during resolution

    @classmethod
    def from_xml(cls, styles_xml: bytes | None) -> StyleResolver:
        """Create a StyleResolver from styles.xml bytes.

        Args:
            styles_xml: Raw bytes of styles.xml, or None if missing.

        Returns:
            StyleResolver instance.
        """
        if styles_xml is None:
            return cls()

        from kaos_office.opc.security import parse_xml_safe

        root = parse_xml_safe(styles_xml)
        styles: dict[str, etree._Element] = {}
        for style_el in root.iter(qn(W, "style")):
            style_id = style_el.get(qn(W, "styleId"))
            if style_id:
                styles[style_id] = style_el

        return cls(_styles=styles)

    def heading_level(self, style_id: str | None) -> int | None:
        """Return heading level (1-6) or None if not a heading.

        Args:
            style_id: The w:pStyle value, or None.

        Returns:
            Heading depth (1-6), or None for non-heading styles.
        """
        if style_id is None:
            return None

        if style_id in self._heading_cache:
            return self._heading_cache[style_id]

        # Start cycle detection for this resolution
        self._seen = set()
        level = self._detect_heading_level(style_id)
        self._heading_cache[style_id] = level
        return level

    def _detect_heading_level(self, style_id: str) -> int | None:
        """Detect heading level via outline level, name, or inheritance."""
        if style_id in self._seen:
            return None  # Cycle detected
        self._seen.add(style_id)

        style_el = self._styles.get(style_id)
        if style_el is None:
            return None

        # 1. Check outline level in pPr
        ppr = style_el.find(W_PPR)
        if ppr is not None:
            outline = ppr.find(W_OUTLINE_LVL)
            if outline is not None:
                val = outline.get(W_VAL)
                if val is not None:
                    try:
                        # outlineLvl is 0-based, heading depth is 1-based
                        return min(int(val) + 1, 6)
                    except ValueError:
                        pass

        # 2. Check style name pattern
        name_el = style_el.find(W_NAME)
        if name_el is not None:
            name = name_el.get(W_VAL) or ""
            m = _HEADING_NAME_RE.match(name)
            if m:
                return min(int(m.group(1)), 6)
            # TOC Heading is not a real heading
            if _TOC_HEADING_RE.match(name):
                return None

        # 3. Walk inheritance chain
        based_on = style_el.find(W_BASED_ON)
        if based_on is not None:
            parent_id = based_on.get(W_VAL)
            if parent_id and parent_id != style_id:
                return self._detect_heading_level(parent_id)

        return None

    def is_code_style(self, style_id: str | None) -> bool:
        """Check if a style ID indicates a code/preformatted block.

        Args:
            style_id: The w:pStyle value.

        Returns:
            True if the style name suggests code.
        """
        if style_id is None:
            return False
        style_el = self._styles.get(style_id)
        if style_el is None:
            return False
        name_el = style_el.find(W_NAME)
        if name_el is None:
            return False
        name = (name_el.get(W_VAL) or "").lower()
        return name in ("code", "source code", "htmlcode", "html code", "macro text")

    def has_style(self, style_id: str) -> bool:
        """Check if a style ID exists in the document."""
        return style_id in self._styles
