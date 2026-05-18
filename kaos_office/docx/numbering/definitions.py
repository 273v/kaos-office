"""Numbering definitions: typed data classes for ``numbering.xml``.

DOCX numbering uses three layers of indirection:

1. ``<w:p>`` carries ``<w:numPr><w:numId/><w:ilvl/></w:numPr>`` — a
   reference to a *numbering instance* at an *indent level*.
2. ``<w:num>`` maps each ``numId`` to an ``abstractNumId`` and may
   override individual levels (``<w:lvlOverride>``, including
   ``<w:startOverride>``).
3. ``<w:abstractNum>`` defines the format for each level
   (``numFmt``, ``lvlText``, ``start``, ``lvlRestart``, ``isLgl``,
   ``suff``).

These dataclasses model that surface. The :class:`NumberingDefinitions`
container resolves ``(num_id, ilvl)`` to an effective
:class:`LevelDefinition`, honoring ``lvlOverride`` and
``startOverride`` so callers see one consistent definition.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class LevelDefinition:
    """Format definition for one ``ilvl`` of an abstract numbering."""

    level: int
    """Indent level, 0-8."""

    num_format: str
    """Word ``numFmt`` value — ``"decimal"``, ``"lowerLetter"``,
    ``"lowerRoman"``, ``"bullet"``, etc."""

    level_text: str
    """``lvlText`` template, e.g. ``"%1."``, ``"(%2)"``, ``"%1(%2)(%3)"``.

    ``%N`` placeholders are 1-based and **absolute**: ``%1`` always
    means level 0, ``%2`` always means level 1, etc., regardless of
    which level this definition belongs to.
    """

    start_value: int = 1
    """Initial counter value (from ``<w:start w:val="..."/>``)."""

    restart_after_level: int | None = None
    """Honored ``<w:lvlRestart>`` — the level that triggers a restart of
    this level's counter when it advances. ``None`` means "no explicit
    restart configured" (default Word behavior: restart on the parent
    level)."""

    is_legal: bool = False
    """``<w:isLgl/>`` — when True, every ``%N`` substitution renders as
    decimal regardless of the referenced level's ``num_format``."""

    suff: str = "tab"
    """``<w:suff>`` — separator between label and body text:
    ``"tab"``, ``"space"``, or ``"nothing"``. Default in Word is
    ``"tab"``."""

    paragraph_style: str | None = None
    """``<w:pStyle>`` — paragraph style linked to this level. Used to
    resolve numbering for paragraphs that carry no inline ``<w:numPr>``
    but inherit numbering through their style."""

    @property
    def is_bullet(self) -> bool:
        """True for bullet/none formats — counter has no visible value."""
        return self.num_format in {"bullet", "none", ""}


@dataclass(frozen=True)
class AbstractNum:
    """Definition of a Word abstract numbering — a reusable template
    referenced by zero or more numbering instances.
    """

    abstract_num_id: str
    levels: dict[int, LevelDefinition] = field(default_factory=dict)
    multi_level_type: str | None = None
    name: str | None = None


@dataclass(frozen=True)
class NumInstance:
    """A specific numbering instance (``<w:num>``).

    Carries optional per-level overrides — most commonly
    ``<w:startOverride w:val="5"/>`` to start numbering at 5 rather
    than the abstractNum's declared ``start_value``.
    """

    num_id: str
    abstract_num_id: str
    level_overrides: dict[int, LevelOverride] = field(default_factory=dict)


@dataclass(frozen=True)
class LevelOverride:
    """``<w:lvlOverride>`` payload — what to override on one level of
    the referenced ``abstractNum``.
    """

    start_override: int | None = None
    """``<w:startOverride>`` — explicit starting counter value."""

    level_definition: LevelDefinition | None = None
    """If the override includes a full ``<w:lvl>`` redefining the
    format at this level, store it here."""


class NumberingDefinitions:
    """Effective numbering definitions for a single DOCX.

    Resolves ``(num_id, ilvl)`` to a :class:`LevelDefinition` that has
    already absorbed any ``<w:lvlOverride>`` from the matching
    ``<w:num>``. Read-only after construction.
    """

    __slots__ = ("_abstract_nums", "_num_instances", "_pstyle_index")

    def __init__(
        self,
        abstract_nums: dict[str, AbstractNum] | None = None,
        num_instances: dict[str, NumInstance] | None = None,
    ) -> None:
        self._abstract_nums: dict[str, AbstractNum] = abstract_nums or {}
        self._num_instances: dict[str, NumInstance] = num_instances or {}
        # Build an index from paragraph style id → (num_id, ilvl) for
        # style-linked numbering resolution. The first num instance
        # whose abstractNum links the style wins; this matches Word's
        # "first match" semantics for ambiguously linked styles.
        self._pstyle_index: dict[str, tuple[str, int]] = {}
        for num_id, inst in self._num_instances.items():
            abstract = self._abstract_nums.get(inst.abstract_num_id)
            if abstract is None:
                continue
            for ilvl, lvl in abstract.levels.items():
                if lvl.paragraph_style is None:
                    continue
                self._pstyle_index.setdefault(lvl.paragraph_style, (num_id, ilvl))

    def has_num_id(self, num_id: str) -> bool:
        """True if ``num_id`` has a known abstract definition."""
        inst = self._num_instances.get(num_id)
        return inst is not None and inst.abstract_num_id in self._abstract_nums

    def get_level_definition(self, num_id: str, ilvl: int) -> LevelDefinition | None:
        """Return the effective level definition, honoring overrides.

        ``None`` when ``num_id`` is unknown, the linked abstract is
        unknown, or the abstract lacks a definition for ``ilvl``.
        """
        inst = self._num_instances.get(num_id)
        if inst is None:
            return None
        abstract = self._abstract_nums.get(inst.abstract_num_id)
        if abstract is None:
            return None

        base = abstract.levels.get(ilvl)
        override = inst.level_overrides.get(ilvl)

        if override is not None and override.level_definition is not None:
            # Full level redefinition wins outright.
            return self._apply_start_override(override.level_definition, override.start_override)

        if base is None:
            return None
        return self._apply_start_override(base, override.start_override if override else None)

    @staticmethod
    def _apply_start_override(
        level: LevelDefinition, start_override: int | None
    ) -> LevelDefinition:
        if start_override is None or start_override == level.start_value:
            return level
        return LevelDefinition(
            level=level.level,
            num_format=level.num_format,
            level_text=level.level_text,
            start_value=start_override,
            restart_after_level=level.restart_after_level,
            is_legal=level.is_legal,
            suff=level.suff,
            paragraph_style=level.paragraph_style,
        )

    def resolve_pstyle(self, style_id: str) -> tuple[str, int] | None:
        """Return ``(num_id, ilvl)`` for a paragraph style that links a
        numbering definition, or ``None`` if no such link exists.

        Used by the reader when a paragraph has no inline
        ``<w:numPr>`` but inherits numbering through its
        ``<w:pStyle>``.
        """
        return self._pstyle_index.get(style_id)
