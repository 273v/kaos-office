"""SmartArt text extraction via OPC fallback.

python-pptx cannot extract SmartArt text (issue #83, open since 2014).
This module parses diagrams/data1.xml directly via the OPC layer to
extract text from SmartArt diagram nodes.
"""

from __future__ import annotations

from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import A_T, DGM_PT, DGM_PT_LST, DGM_REL_IDS, R, qn
from kaos_office.opc.package import OPCPackage
from kaos_office.opc.relationships import RelationshipManager


def extract_smartart_texts(
    graphic_data_el: etree._Element,
    pkg: OPCPackage,
    slide_rels: RelationshipManager,
    slide_part_dir: str,
) -> list[str]:
    """Extract text strings from a SmartArt diagram via OPC fallback.

    Args:
        graphic_data_el: The <a:graphicData> element containing dgm:relIds.
        pkg: The open OPC package.
        slide_rels: Relationship manager for the slide part.
        slide_part_dir: Directory of the slide part (e.g., "ppt/slides").

    Returns:
        List of text strings from SmartArt nodes, in document order.
    """
    # Find dgm:relIds element
    rel_ids_el = graphic_data_el.find(f".//{DGM_REL_IDS}")
    if rel_ids_el is None:
        return []

    # Get the data model relationship ID (r:dm)
    dm_rid = rel_ids_el.get(qn(R, "dm"))
    if not dm_rid:
        return []

    # Resolve the relationship to a part path
    dm_target = slide_rels.resolve(dm_rid)
    if not dm_target:
        return []

    # Build full part path relative to package root
    if dm_target.startswith("/"):
        dm_part = dm_target.lstrip("/")
    elif dm_target.startswith(".."):
        # Relative path from slide directory
        parts = slide_part_dir.rstrip("/").split("/")
        for seg in dm_target.split("/"):
            if seg == "..":
                parts.pop()
            else:
                parts.append(seg)
        dm_part = "/".join(parts)
    else:
        dm_part = f"{slide_part_dir}/{dm_target}"

    if not pkg.has_part(dm_part):
        return []

    try:
        data_xml = pkg.read_xml(dm_part)
    except Exception:
        return []

    return _extract_texts_from_data_model(data_xml)


def _extract_texts_from_data_model(root: etree._Element) -> list[str]:
    """Extract text from dgm:pt nodes in a SmartArt data model.

    Iterates dgm:ptLst → dgm:pt where type != "doc", extracts a:t text.
    """
    texts: list[str] = []

    pt_lst = root.find(f".//{DGM_PT_LST}")
    if pt_lst is None:
        # Try iterating directly if ptLst is the root or nested differently
        for pt in root.iter(DGM_PT):
            _collect_pt_text(pt, texts)
        return texts

    for pt in pt_lst.findall(DGM_PT):
        _collect_pt_text(pt, texts)

    return texts


def _collect_pt_text(pt: etree._Element, texts: list[str]) -> None:
    """Collect text from a single dgm:pt element if it's a content node."""
    pt_type = pt.get("type", "node")
    # Skip document root nodes and presentation/transition nodes
    if pt_type in ("doc", "pres", "sibTrans", "parTrans", "asst"):
        return

    # Collect all a:t text within this point
    parts: list[str] = []
    for t_el in pt.iter(A_T):
        if t_el.text:
            parts.append(t_el.text)

    combined = " ".join(parts).strip()
    if combined:
        texts.append(combined)
