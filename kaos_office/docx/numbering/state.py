"""Running counter state machine for DOCX list numbering.

The heart of attorney-grade citation: as paragraphs stream through the
DOCX reader, this class tracks the running counter for each
``num_id`` at each ``ilvl`` and emits the rendered numbering label
(``"11."``, ``"(a)"``, ``"11(a)(i)"``, etc.) for each numbered
paragraph.

Algorithm follows the canonical kelvin-office implementation with two
corrections:

* ``lvlRestart`` is enforced rather than only parsed.
* ``startOverride`` (from ``<w:lvlOverride><w:startOverride/>``) is
  applied through :class:`NumberingDefinitions.get_level_definition`
  rather than ignored.
"""

from __future__ import annotations

import re
from typing import Final

from kaos_office.docx.numbering.definitions import (
    LevelDefinition,
    NumberingDefinitions,
)
from kaos_office.docx.numbering.formatters import format_number

_PLACEHOLDER_RE: Final[re.Pattern[str]] = re.compile(r"%([1-9])")


class NumberingState:
    """Running counter machine for one DOCX document.

    State is keyed by ``num_id``. ``last_num_id`` / ``last_level`` are
    tracked so we know when the level transitioned (deeper / shallower
    / same) — the case analysis in :meth:`_update_counters` mirrors
    Word's renderer.

    Not thread-safe. Construct one per document and feed paragraphs in
    document order.
    """

    __slots__ = ("_counters", "_definitions", "_last_level", "_last_num_id")

    def __init__(self, definitions: NumberingDefinitions) -> None:
        self._definitions = definitions
        self._counters: dict[str, dict[int, int]] = {}
        self._last_num_id: str | None = None
        self._last_level: int | None = None

    @property
    def definitions(self) -> NumberingDefinitions:
        return self._definitions

    def get_formatted_label(self, num_id: str, level: int) -> str:
        """Advance the counter for ``(num_id, level)`` and return the
        rendered numbering label.

        Returns ``""`` when the level is undefined or maps to a bullet
        with no visible numeral (see :func:`format_number`). The empty
        string is the signal to callers that there is no attorney-grade
        citation token for this paragraph (still emit the paragraph,
        just without a label).
        """
        level_def = self._definitions.get_level_definition(num_id, level)
        if level_def is None:
            return ""

        self._update_counters(num_id, level, level_def)
        label = self._format_label(level_def.level_text, num_id, level_def)

        if not label:
            # Bullet / none format → keep the format's literal glyph if
            # any. format_number on the current level's counter handles
            # bullet/"none" correctly.
            label = format_number(self._counters[num_id].get(level, 1), level_def.num_format)
        return label

    # ── Counter state ──────────────────────────────────────────────

    def _update_counters(self, num_id: str, level: int, level_def: LevelDefinition) -> None:
        bucket = self._counters.setdefault(num_id, {})
        same_flow = num_id == self._last_num_id and self._last_level is not None

        # lvlRestart applies within a single numId flow. When an
        # unrelated numId interrupts (e.g. a heading numbered "Section
        # 1." appearing between two items of a different sub-clause
        # list), each numId's counters survive untouched — that's the
        # contract real Word documents rely on for two-axis
        # heading / sub-clause numbering.
        if same_flow and self._last_level is not None and level <= self._last_level:
            self._restart_deeper_levels(num_id, level)

        if level in bucket:
            bucket[level] += 1
        else:
            self._init_level(num_id, level, level_def)

        self._last_num_id = num_id
        self._last_level = level

    def _init_level(self, num_id: str, level: int, level_def: LevelDefinition) -> None:
        bucket = self._counters.setdefault(num_id, {})
        bucket[level] = level_def.start_value

    def _restart_deeper_levels(self, num_id: str, level: int) -> None:
        """Reset deeper levels whose ``lvlRestart`` policy requires it
        when level ``level`` advances.

        Word's renderer restarts a child counter whenever a configured
        ancestor advances. Without this, a sub-clause counter would
        continue past the visible section break — producing
        ``Section 2(c)`` after ``Section 1(b)`` instead of
        ``Section 2(a)``.

        ``restart_after_level`` (from ``<w:lvlRestart w:val="..."/>``)
        semantics:

        * ``None`` — default Word behavior: restart whenever ANY
          shallower level advances.
        * ``0`` — Word's "never restart" sentinel: keep the counter
          regardless of shallower advances.
        * ``k`` (1-based) — restart when level ``k - 1`` (0-based) or
          any shallower level advances. So ``lvlRestart=2`` on a
          level-3 counter means "restart when level 1 advances."
        """
        bucket = self._counters.get(num_id)
        if not bucket:
            return
        for deeper_level in list(bucket.keys()):
            if deeper_level <= level:
                continue
            deeper_def = self._definitions.get_level_definition(num_id, deeper_level)
            if deeper_def is None:
                # Unknown level definition — fall back to default
                # restart behavior (any shallower advance restarts).
                del bucket[deeper_level]
                continue
            restart_after = deeper_def.restart_after_level
            if restart_after is None:
                del bucket[deeper_level]
                continue
            if restart_after == 0:
                # "Never restart" — keep counter untouched.
                continue
            # Word's @w:val is a 1-based level reference; the deeper
            # level restarts only when a level at or shallower than
            # (restart_after - 1) advances.
            restart_threshold = restart_after - 1
            if level <= restart_threshold:
                del bucket[deeper_level]

    # ── Label rendering ─────────────────────────────────────────────

    def _format_label(self, pattern: str, num_id: str, current_level_def: LevelDefinition) -> str:
        if not pattern:
            return ""

        def _replace(match: re.Match[str]) -> str:
            placeholder = int(match.group(1))
            # `%N` is absolute and 1-based: %1 → level 0, %2 → level 1.
            referenced_level = placeholder - 1
            counter = self._counters.get(num_id, {}).get(referenced_level)
            if counter is None:
                return ""
            referenced_def = self._definitions.get_level_definition(num_id, referenced_level)
            # isLgl forces decimal for every referenced level.
            if current_level_def.is_legal:
                return format_number(counter, "decimal")
            if referenced_def is None:
                return format_number(counter, current_level_def.num_format)
            return format_number(counter, referenced_def.num_format)

        return _PLACEHOLDER_RE.sub(_replace, pattern)
