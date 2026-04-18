"""OPC Relationship Manager.

Parses .rels files and provides bidirectional lookup:
- By relationship ID (rId → target path)
- By relationship type (URI → list of relationships)
"""

from __future__ import annotations

from dataclasses import dataclass

from kaos_office.opc.security import parse_xml_safe

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True, slots=True)
class Relationship:
    """A single OPC relationship."""

    id: str  # rId1, rId2, ...
    type: str  # Full URI
    target: str  # Relative path (e.g., styles.xml, media/image1.png)
    external: bool = False  # True for TargetMode="External"


class RelationshipManager:
    """Manages OPC relationships for a single source part.

    Bidirectional: lookup by rId (for reading) or by type (for discovery).
    """

    def __init__(self) -> None:
        self._by_id: dict[str, Relationship] = {}
        self._by_type: dict[str, list[Relationship]] = {}

    @classmethod
    def parse(cls, rels_xml: bytes) -> RelationshipManager:
        """Parse a .rels XML file into a RelationshipManager.

        Args:
            rels_xml: Raw XML bytes of the .rels file.

        Returns:
            Populated RelationshipManager.
        """
        mgr = cls()
        root = parse_xml_safe(rels_xml)

        for child in root:
            tag = child.tag
            # Handle both namespaced and non-namespaced
            if not tag.endswith("Relationship"):
                continue

            rel_id = child.get("Id", "")
            rel_type = child.get("Type", "")
            target = child.get("Target", "")
            target_mode = child.get("TargetMode", "")
            external = target_mode.lower() == "external"

            if rel_id and rel_type:
                rel = Relationship(
                    id=rel_id,
                    type=rel_type,
                    target=target,
                    external=external,
                )
                mgr._by_id[rel_id] = rel
                mgr._by_type.setdefault(rel_type, []).append(rel)

        return mgr

    def resolve(self, rel_id: str) -> str | None:
        """Resolve a relationship ID to a target path.

        Args:
            rel_id: The relationship ID (e.g., "rId7").

        Returns:
            Target path string, or None if not found.
        """
        rel = self._by_id.get(rel_id)
        return rel.target if rel else None

    def get(self, rel_id: str) -> Relationship | None:
        """Get a relationship by ID.

        Args:
            rel_id: The relationship ID.

        Returns:
            Relationship object, or None if not found.
        """
        return self._by_id.get(rel_id)

    def by_type(self, rel_type: str) -> list[Relationship]:
        """Find all relationships of a given type.

        Args:
            rel_type: Full relationship type URI.

        Returns:
            List of matching relationships (may be empty).
        """
        return list(self._by_type.get(rel_type, []))

    def first_target(self, rel_type: str) -> str | None:
        """Get the target of the first relationship of a given type.

        Args:
            rel_type: Full relationship type URI.

        Returns:
            Target path string, or None if no matching relationship.
        """
        rels = self._by_type.get(rel_type, [])
        return rels[0].target if rels else None

    # --- L2 Write Methods ---

    def add(
        self,
        rel_type: str,
        target: str,
        *,
        rel_id: str | None = None,
        external: bool = False,
    ) -> Relationship:
        """Add a new relationship. Auto-generates ID if not provided."""
        if rel_id is None:
            rel_id = self.next_id()
        rel = Relationship(id=rel_id, type=rel_type, target=target, external=external)
        self._by_id[rel_id] = rel
        self._by_type.setdefault(rel_type, []).append(rel)
        return rel

    def next_id(self) -> str:
        """Generate the next available rId (rId1, rId2, ...)."""
        n = 1
        while f"rId{n}" in self._by_id:
            n += 1
        return f"rId{n}"

    def serialize(self) -> bytes:
        """Serialize to .rels XML bytes."""
        from lxml import etree  # ty: ignore[unresolved-import]

        root = etree.Element("Relationships", nsmap={None: _RELS_NS})

        for rel in sorted(self._by_id.values(), key=lambda r: r.id):
            attribs: dict[str, str] = {
                "Id": rel.id,
                "Type": rel.type,
                "Target": rel.target,
            }
            if rel.external:
                attribs["TargetMode"] = "External"
            etree.SubElement(root, "Relationship", **attribs)

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def __len__(self) -> int:
        return len(self._by_id)

    def __contains__(self, rel_id: str) -> bool:
        return rel_id in self._by_id
