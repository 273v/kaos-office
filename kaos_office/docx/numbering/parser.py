"""Parser for Word ``numbering.xml`` → :class:`NumberingDefinitions`."""

from __future__ import annotations

from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.docx.numbering.definitions import (
    AbstractNum,
    LevelDefinition,
    LevelOverride,
    NumberingDefinitions,
    NumInstance,
)
from kaos_office.ooxml.namespace import (
    W_ABSTRACT_NUM,
    W_ABSTRACT_NUM_ID,
    W_IS_LGL,
    W_LVL,
    W_LVL_OVERRIDE,
    W_LVL_RESTART,
    W_LVL_TEXT,
    W_NUM,
    W_NUM_FMT,
    W_PSTYLE,
    W_START,
    W_START_OVERRIDE,
    W_SUFF,
    W_VAL,
    W,
    qn,
)


def parse_numbering_xml(numbering_xml: bytes | None) -> NumberingDefinitions:
    """Parse ``word/numbering.xml`` into a :class:`NumberingDefinitions`.

    Returns an empty (but valid) :class:`NumberingDefinitions` when
    ``numbering_xml`` is ``None`` or empty. Uses the package's
    safe-XML parser (``parse_xml_safe``) — never raw ``etree.parse``.
    """
    if not numbering_xml:
        return NumberingDefinitions()

    # Local import: parse_xml_safe pulls in lxml + the package security
    # policy. Keeping the import lazy so callers that never hit numbered
    # documents do not pay the import cost.
    from kaos_office.opc.security import parse_xml_safe

    root = parse_xml_safe(numbering_xml)

    abstract_nums: dict[str, AbstractNum] = {}
    for an_el in root.iter(W_ABSTRACT_NUM):
        an = _parse_abstract_num(an_el)
        if an is not None:
            abstract_nums[an.abstract_num_id] = an

    num_instances: dict[str, NumInstance] = {}
    for num_el in root.iter(W_NUM):
        inst = _parse_num_instance(num_el)
        if inst is not None:
            num_instances[inst.num_id] = inst

    return NumberingDefinitions(
        abstract_nums=abstract_nums,
        num_instances=num_instances,
    )


def _parse_abstract_num(elem: etree._Element) -> AbstractNum | None:
    abstract_id = elem.get(W_ABSTRACT_NUM_ID)
    if abstract_id is None:
        return None
    levels: dict[int, LevelDefinition] = {}
    for lvl_el in elem.findall(W_LVL):
        lvl = _parse_level(lvl_el)
        if lvl is not None:
            levels[lvl.level] = lvl
    multi_level_type_el = elem.find(qn(W, "multiLevelType"))
    multi_level_type = multi_level_type_el.get(W_VAL) if multi_level_type_el is not None else None
    name_el = elem.find(qn(W, "name"))
    name = name_el.get(W_VAL) if name_el is not None else None
    return AbstractNum(
        abstract_num_id=abstract_id,
        levels=levels,
        multi_level_type=multi_level_type,
        name=name,
    )


def _parse_level(elem: etree._Element) -> LevelDefinition | None:
    ilvl_raw = elem.get(qn(W, "ilvl"))
    if ilvl_raw is None:
        return None
    try:
        level = int(ilvl_raw)
    except ValueError:
        return None

    num_fmt_el = elem.find(W_NUM_FMT)
    num_format = num_fmt_el.get(W_VAL) if num_fmt_el is not None else "decimal"
    if num_format is None:
        num_format = "decimal"

    lvl_text_el = elem.find(W_LVL_TEXT)
    level_text = lvl_text_el.get(W_VAL) if lvl_text_el is not None else ""
    if level_text is None:
        level_text = ""

    start_el = elem.find(W_START)
    start_value = _parse_int(start_el.get(W_VAL) if start_el is not None else None, default=1)
    if start_value is None:
        # _parse_int returns None when an out-of-band default is None;
        # we always pass default=1 here, so this is a defensive fallback.
        start_value = 1

    lvl_restart_el = elem.find(W_LVL_RESTART)
    restart_after_level = _parse_int(
        lvl_restart_el.get(W_VAL) if lvl_restart_el is not None else None,
        default=None,
    )

    # <w:isLgl/> is a presence-only flag (no @w:val required).
    is_legal = elem.find(W_IS_LGL) is not None

    suff_el = elem.find(W_SUFF)
    suff = suff_el.get(W_VAL) if suff_el is not None else "tab"
    if suff is None:
        suff = "tab"

    pstyle_el = elem.find(W_PSTYLE)
    paragraph_style = pstyle_el.get(W_VAL) if pstyle_el is not None else None

    return LevelDefinition(
        level=level,
        num_format=num_format,
        level_text=level_text,
        start_value=start_value,
        restart_after_level=restart_after_level,
        is_legal=is_legal,
        suff=suff,
        paragraph_style=paragraph_style,
    )


def _parse_num_instance(elem: etree._Element) -> NumInstance | None:
    num_id = elem.get(qn(W, "numId"))
    if num_id is None:
        return None
    abstract_ref = elem.find(W_ABSTRACT_NUM_ID)
    if abstract_ref is None:
        return None
    abstract_num_id = abstract_ref.get(W_VAL) or ""
    if not abstract_num_id:
        return None

    overrides: dict[int, LevelOverride] = {}
    for lo_el in elem.findall(W_LVL_OVERRIDE):
        override = _parse_level_override(lo_el)
        if override is None:
            continue
        ilvl, payload = override
        overrides[ilvl] = payload

    return NumInstance(
        num_id=num_id,
        abstract_num_id=abstract_num_id,
        level_overrides=overrides,
    )


def _parse_level_override(
    elem: etree._Element,
) -> tuple[int, LevelOverride] | None:
    ilvl_raw = elem.get(qn(W, "ilvl"))
    if ilvl_raw is None:
        return None
    try:
        ilvl = int(ilvl_raw)
    except ValueError:
        return None

    start_override_el = elem.find(W_START_OVERRIDE)
    start_override = _parse_int(
        start_override_el.get(W_VAL) if start_override_el is not None else None,
        default=None,
    )

    inner_lvl_el = elem.find(W_LVL)
    level_definition = _parse_level(inner_lvl_el) if inner_lvl_el is not None else None

    return ilvl, LevelOverride(
        start_override=start_override,
        level_definition=level_definition,
    )


def _parse_int(raw: str | None, *, default: int | None) -> int | None:
    if raw is None or raw == "":
        return default
    try:
        return int(raw)
    except ValueError:
        return default
