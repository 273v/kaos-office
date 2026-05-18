"""Tests for ``NumberingState`` — the running-counter machine.

The cases that bit kelvin-office (and that every legal-docs reader
trips on if these regress):

* ``%N`` placeholders are absolute level references, not relative.
* Letter / roman wrap-around at level boundaries.
* ``lvlRestart`` actually restarts deeper levels when the configured
  shallower level advances.
* ``<w:lvlOverride><w:startOverride/>`` shifts the starting counter.
* ``isLgl`` forces decimal at every ``%N`` substitution.
"""

from __future__ import annotations

import pytest

from kaos_office.docx.numbering import (
    NumberingState,
    parse_numbering_xml,
)
from kaos_office.ooxml.namespace import W


def _numbering_xml(*defs: str) -> bytes:
    return f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  {"".join(defs)}
</w:numbering>""".encode()


# Three-level decimal / lowerLetter / lowerRoman pattern — the textbook
# legal section pattern ``Section 11(a)(i)``.
_LEGAL_THREE_LEVEL = _numbering_xml(
    """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="%1."/>
  </w:lvl>
  <w:lvl w:ilvl="1">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerLetter"/>
    <w:lvlText w:val="(%2)"/>
  </w:lvl>
  <w:lvl w:ilvl="2">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerRoman"/>
    <w:lvlText w:val="(%3)"/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
)


class TestSingleLevelDecimal:
    def setup_method(self) -> None:
        self.defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        self.state = NumberingState(self.defs)

    def test_first_emits_one(self) -> None:
        assert self.state.get_formatted_label("1", 0) == "1."

    def test_running_counter(self) -> None:
        labels = [self.state.get_formatted_label("1", 0) for _ in range(3)]
        assert labels == ["1.", "2.", "3."]


class TestLowerLetterLevel:
    """``%2`` at level 1 references its own counter — must be letters,
    not the parent's decimal counter."""

    def setup_method(self) -> None:
        self.defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        self.state = NumberingState(self.defs)
        self.state.get_formatted_label("1", 0)  # establish level 0

    def test_descend_to_level_1(self) -> None:
        labels = [self.state.get_formatted_label("1", 1) for _ in range(3)]
        assert labels == ["(a)", "(b)", "(c)"]


class TestLowerRomanLevel:
    def setup_method(self) -> None:
        self.defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        self.state = NumberingState(self.defs)
        self.state.get_formatted_label("1", 0)
        self.state.get_formatted_label("1", 1)

    def test_descend_to_level_2(self) -> None:
        labels = [self.state.get_formatted_label("1", 2) for _ in range(4)]
        assert labels == ["(i)", "(ii)", "(iii)", "(iv)"]


class TestFullThreeLevelSequence:
    """The whole "Section 11(a)(i)" pattern in one go — the canonical
    attorney citation scenario from the 2026-05-18 NDA case."""

    def test_section_eleven_a_i(self) -> None:
        defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        state = NumberingState(defs)
        # 10 prior section-0 paragraphs, then level-1 / level-2 deep dive.
        for _ in range(11):
            state.get_formatted_label("1", 0)
        # We are now at "Section 11."
        # Descend into (a)
        a_label = state.get_formatted_label("1", 1)
        # Descend into (i)
        i_label = state.get_formatted_label("1", 2)
        assert a_label == "(a)"
        assert i_label == "(i)"

    def test_counter_at_level_zero_advances_to_eleven(self) -> None:
        defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        state = NumberingState(defs)
        last = ""
        for _ in range(11):
            last = state.get_formatted_label("1", 0)
        assert last == "11."


class TestRestartOnShallowerAdvance:
    """When section 1 has (a), (b), (c) and we jump back to level 0
    for section 2, the (a) counter must reset to ``a`` — otherwise an
    attorney citing ``Section 2(a)`` would actually point at the fourth
    sub-clause.
    """

    def test_restart_letters_after_shallow_advance(self) -> None:
        defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        state = NumberingState(defs)
        state.get_formatted_label("1", 0)  # 1.
        state.get_formatted_label("1", 1)  # (a)
        state.get_formatted_label("1", 1)  # (b)
        state.get_formatted_label("1", 1)  # (c)
        # Jump back to a new section
        section_two = state.get_formatted_label("1", 0)
        next_subclause = state.get_formatted_label("1", 1)
        assert section_two == "2."
        assert next_subclause == "(a)"

    def test_restart_clears_grandchildren_too(self) -> None:
        defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        state = NumberingState(defs)
        state.get_formatted_label("1", 0)  # 1.
        state.get_formatted_label("1", 1)  # (a)
        state.get_formatted_label("1", 2)  # (i)
        state.get_formatted_label("1", 2)  # (ii)
        state.get_formatted_label("1", 0)  # 2.
        state.get_formatted_label("1", 1)  # (a) again
        next_roman = state.get_formatted_label("1", 2)
        assert next_roman == "(i)"


class TestLvlRestartConfigured:
    """When the abstractNum sets ``<w:lvlRestart w:val="0"/>`` on level
    1 (here Word's semantics for "do not restart"), descending back
    into level 1 after returning to level 0 keeps counting.

    Note: Word's @w:val=0 on lvlRestart means "do not restart." That's
    the case being asserted here.
    """

    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="%1."/>
  </w:lvl>
  <w:lvl w:ilvl="1">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerLetter"/>
    <w:lvlText w:val="(%2)"/>
    <w:lvlRestart w:val="0"/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )

    def test_no_restart_keeps_counter(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        state.get_formatted_label("1", 0)  # 1.
        state.get_formatted_label("1", 1)  # (a)
        state.get_formatted_label("1", 1)  # (b)
        state.get_formatted_label("1", 0)  # 2.
        # With lvlRestart=0 we expect the letter counter to persist.
        next_letter = state.get_formatted_label("1", 1)
        assert next_letter == "(c)"


class TestStartOverride:
    """``<w:startOverride w:val="5"/>`` shifts the starting counter for
    a specific numId. Common when a contract has explicit "this
    Schedule starts at Section 5" semantics.
    """

    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="%1."/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="5">
  <w:abstractNumId w:val="0"/>
  <w:lvlOverride w:ilvl="0">
    <w:startOverride w:val="5"/>
  </w:lvlOverride>
</w:num>
"""
    )

    def test_starts_at_five(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        first = state.get_formatted_label("5", 0)
        second = state.get_formatted_label("5", 0)
        assert first == "5."
        assert second == "6."


class TestIsLglForcesDecimal:
    """``<w:isLgl/>`` on a level renders every ``%N`` substitution as
    decimal even when the referenced level's ``numFmt`` is e.g.
    lowerLetter. Produces ``1.1.1`` style legal outlines.
    """

    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="%1."/>
  </w:lvl>
  <w:lvl w:ilvl="1">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerLetter"/>
    <w:lvlText w:val="%1.%2"/>
    <w:isLgl/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )

    def test_isLgl_forces_decimal_at_referenced_levels(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        state.get_formatted_label("1", 0)
        label = state.get_formatted_label("1", 1)
        # Without isLgl, %2 would render as "a" (lowerLetter).
        # With isLgl, both %1 and %2 are forced to decimal.
        assert label == "1.1"


class TestBulletLevel:
    """Bullet-format levels return the bullet glyph, not an empty string."""

    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:numFmt w:val="bullet"/>
    <w:lvlText w:val=""/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )

    def test_returns_bullet_glyph(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        label = state.get_formatted_label("1", 0)
        assert label == "•"


class TestUnknownNumId:
    def test_unknown_num_id_returns_empty_string(self) -> None:
        defs = parse_numbering_xml(_LEGAL_THREE_LEVEL)
        state = NumberingState(defs)
        assert state.get_formatted_label("999", 0) == ""


class TestStartValueRespected:
    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="7"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="Section %1."/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )

    def test_start_value_seven(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        first = state.get_formatted_label("1", 0)
        assert first == "Section 7."


class TestMultilineLvlTextTemplate:
    """``lvlText="%1(%2)(%3)"`` at level 2 produces ``11(a)(i)``
    directly — the canonical attorney citation token."""

    XML = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="%1"/>
  </w:lvl>
  <w:lvl w:ilvl="1">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerLetter"/>
    <w:lvlText w:val="%1(%2)"/>
  </w:lvl>
  <w:lvl w:ilvl="2">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerRoman"/>
    <w:lvlText w:val="%1(%2)(%3)"/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )

    def test_section_eleven_a_i_compact(self) -> None:
        defs = parse_numbering_xml(self.XML)
        state = NumberingState(defs)
        for _ in range(11):
            state.get_formatted_label("1", 0)
        state.get_formatted_label("1", 1)  # 11(a)
        i = state.get_formatted_label("1", 2)  # 11(a)(i)
        assert i == "11(a)(i)"


@pytest.mark.parametrize(
    ("counter", "expected"),
    [(1, "Section 1"), (11, "Section 11"), (26, "Section 26"), (27, "Section 27")],
)
def test_section_decimal_label(counter: int, expected: str) -> None:
    """Sanity check: the decimal formatter inside the state machine
    survives boundary values for a typical ``"Section %1"`` template."""
    xml = _numbering_xml(
        """\
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:lvlText w:val="Section %1"/>
  </w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
"""
    )
    defs = parse_numbering_xml(xml)
    state = NumberingState(defs)
    label = ""
    for _ in range(counter):
        label = state.get_formatted_label("1", 0)
    assert label == expected
