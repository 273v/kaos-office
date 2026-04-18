"""OPC Content Types map.

Parses [Content_Types].xml and maps part names/extensions to MIME types.
"""

from __future__ import annotations

from pathlib import PurePosixPath

from kaos_office.opc.security import parse_xml_safe


class ContentTypeMap:
    """Maps part names and extensions to MIME content types.

    Two lookup mechanisms per the OPC spec:
    - Default: file extension → content type
    - Override: specific part name → content type (takes precedence)
    """

    def __init__(self) -> None:
        self._defaults: dict[str, str] = {}  # extension (no dot) → content type
        self._overrides: dict[str, str] = {}  # part name → content type

    @classmethod
    def parse(cls, content_types_xml: bytes) -> ContentTypeMap:
        """Parse [Content_Types].xml bytes into a ContentTypeMap.

        Args:
            content_types_xml: Raw XML bytes of [Content_Types].xml.

        Returns:
            Populated ContentTypeMap.
        """
        ct_map = cls()
        root = parse_xml_safe(content_types_xml)

        for child in root:
            tag = _local_name(child.tag)
            if tag == "Default":
                ext = child.get("Extension", "")
                ctype = child.get("ContentType", "")
                if ext and ctype:
                    ct_map._defaults[ext.lower()] = ctype
            elif tag == "Override":
                part_name = child.get("PartName", "")
                ctype = child.get("ContentType", "")
                if part_name and ctype:
                    # Normalize: ensure leading /
                    if not part_name.startswith("/"):
                        part_name = "/" + part_name
                    ct_map._overrides[part_name] = ctype

        return ct_map

    def get(self, part_name: str) -> str | None:
        """Get the content type for a part.

        Override entries take precedence over default (extension-based) entries.

        Args:
            part_name: The part name (e.g., "word/document.xml").

        Returns:
            MIME content type string, or None if unknown.
        """
        # Normalize: ensure leading /
        normalized = part_name if part_name.startswith("/") else "/" + part_name
        if normalized in self._overrides:
            return self._overrides[normalized]

        # Fall back to extension-based default
        ext = PurePosixPath(part_name).suffix.lstrip(".")
        if ext:
            return self._defaults.get(ext.lower())

        return None

    # --- L2 Write Methods ---

    def add_default(self, extension: str, content_type: str) -> None:
        """Add an extension-based default content type."""
        self._defaults[extension.lower().lstrip(".")] = content_type

    def add_override(self, part_name: str, content_type: str) -> None:
        """Add a part-specific content type override."""
        normalized = part_name if part_name.startswith("/") else "/" + part_name
        self._overrides[normalized] = content_type

    def serialize(self) -> bytes:
        """Serialize to [Content_Types].xml bytes."""
        from lxml import etree  # ty: ignore[unresolved-import]

        ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        root = etree.Element(f"{{{ns}}}Types", nsmap={None: ns})

        for ext, ctype in sorted(self._defaults.items()):
            etree.SubElement(root, "Default", Extension=ext, ContentType=ctype)

        for part_name, ctype in sorted(self._overrides.items()):
            etree.SubElement(root, "Override", PartName=part_name, ContentType=ctype)

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    @property
    def defaults(self) -> dict[str, str]:
        """Extension → content type defaults (read-only copy)."""
        return dict(self._defaults)

    @property
    def overrides(self) -> dict[str, str]:
        """Part name → content type overrides (read-only copy)."""
        return dict(self._overrides)


def _local_name(tag: str) -> str:
    """Strip namespace prefix from a Clark notation tag."""
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag
