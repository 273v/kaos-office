"""OPC Package — the foundation layer for all Office formats.

Wraps a ZIP archive with OPC semantics: parts, relationships, content types,
and security validation. Format-agnostic — same code for DOCX, XLSX, PPTX.

Designed for L1 (read) with extension points for L2 (write) and L3 (round-trip).
"""

from __future__ import annotations

import zipfile
from pathlib import Path, PurePosixPath
from typing import Self

from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.opc.content_types import ContentTypeMap
from kaos_office.opc.relationships import RelationshipManager
from kaos_office.opc.security import parse_xml_safe, validate_zip_security


class OPCPackageError(Exception):
    """Raised when an OPC package cannot be opened or read."""


class OPCPackage:
    """Open Packaging Conventions container (ZIP-based).

    Provides read access to parts, relationships, and content types.
    Designed to support future write/modify operations.

    Usage:
        with OPCPackage.open(path) as pkg:
            doc_xml = pkg.read_xml("word/document.xml")
            styles_bytes = pkg.read_part("word/styles.xml")
            rels = pkg.relationships("word/document.xml")
    """

    def __init__(self, path: str | Path) -> None:
        self._path = Path(path)
        self._zf: zipfile.ZipFile | None = None
        self._content_types: ContentTypeMap | None = None
        self._rels_cache: dict[str, RelationshipManager] = {}
        self._parts_cache: dict[str, bytes] = {}

    @classmethod
    def open(cls, path: str | Path) -> OPCPackage:
        """Open an OPC package for reading.

        Args:
            path: Path to the OPC package file (.docx, .xlsx, .pptx).

        Returns:
            Open OPCPackage instance (use as context manager).

        Raises:
            OPCPackageError: If the file is not a valid ZIP or OPC package.
            OPCSecurityError: If security validation fails.
        """
        pkg = cls(path)
        pkg._open()
        return pkg

    def _open(self) -> None:
        """Open the package for reading."""
        if not self._path.exists():
            raise OPCPackageError(
                f"File not found: {self._path}. "
                "Verify the file path is correct and the file exists."
            )

        try:
            self._zf = zipfile.ZipFile(self._path, "r")
        except zipfile.BadZipFile as exc:
            raise OPCPackageError(
                f"Not a valid ZIP file: {self._path}. "
                "The file may be corrupted or not an Office document."
            ) from exc

        # Security validation
        validate_zip_security(self._zf, self._path.stat().st_size)

        # Parse content types
        if "[Content_Types].xml" not in self._zf.namelist():
            raise OPCPackageError(
                f"Missing [Content_Types].xml in {self._path}. This is not a valid OPC package."
            )
        ct_bytes = self._zf.read("[Content_Types].xml")
        self._content_types = ContentTypeMap.parse(ct_bytes)

    def close(self) -> None:
        """Close the package."""
        if self._zf:
            self._zf.close()
            self._zf = None
        self._parts_cache.clear()
        self._rels_cache.clear()

    def __enter__(self) -> Self:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()

    def read_part(self, part_name: str) -> bytes:
        """Read a part's bytes. Cached after first read.

        Args:
            part_name: Part path within the ZIP (e.g., "word/document.xml").

        Returns:
            Raw bytes of the part.

        Raises:
            OPCPackageError: If the part does not exist.
        """
        if part_name in self._parts_cache:
            return self._parts_cache[part_name]

        self._ensure_open()
        assert self._zf is not None

        # Normalize: strip leading /
        normalized = part_name.lstrip("/")
        if normalized not in self._zf.namelist():
            raise OPCPackageError(
                f"Part not found: '{part_name}' in {self._path}. "
                f"Available parts: {', '.join(sorted(self._zf.namelist())[:10])}..."
            )

        data = self._zf.read(normalized)
        self._parts_cache[part_name] = data
        return data

    def read_xml(self, part_name: str) -> etree._Element:
        """Read and parse a part as XML using secure parser.

        Args:
            part_name: Part path within the ZIP.

        Returns:
            Parsed XML root element.

        Raises:
            OPCPackageError: If the part doesn't exist or XML is malformed.
        """
        data = self.read_part(part_name)
        try:
            return parse_xml_safe(data)
        except etree.XMLSyntaxError as exc:
            raise OPCPackageError(
                f"Malformed XML in part '{part_name}': {exc}. The document may be corrupted."
            ) from exc

    def has_part(self, part_name: str) -> bool:
        """Check if a part exists in the package.

        Args:
            part_name: Part path within the ZIP.

        Returns:
            True if the part exists.
        """
        self._ensure_open()
        assert self._zf is not None
        normalized = part_name.lstrip("/")
        return normalized in self._zf.namelist()

    def list_parts(self) -> list[str]:
        """List all part names in the package.

        Returns:
            Sorted list of part paths.
        """
        self._ensure_open()
        assert self._zf is not None
        return sorted(self._zf.namelist())

    def relationships(self, source: str = "/") -> RelationshipManager:
        """Get the relationship manager for a part.

        Args:
            source: Source part path. Use "/" for root relationships.

        Returns:
            RelationshipManager for the source part.
        """
        if source in self._rels_cache:
            return self._rels_cache[source]

        self._ensure_open()
        assert self._zf is not None

        # Compute .rels path
        rels_path = _rels_path_for(source)
        normalized = rels_path.lstrip("/")

        if normalized in self._zf.namelist():
            rels_xml = self._zf.read(normalized)
            mgr = RelationshipManager.parse(rels_xml)
        else:
            mgr = RelationshipManager()

        self._rels_cache[source] = mgr
        return mgr

    def content_type(self, part_name: str) -> str | None:
        """Get the MIME content type for a part.

        Args:
            part_name: Part path within the ZIP.

        Returns:
            Content type string, or None if unknown.
        """
        if self._content_types is None:
            return None
        return self._content_types.get(part_name)

    @property
    def content_types(self) -> ContentTypeMap | None:
        """The content type map for this package."""
        return self._content_types

    @property
    def path(self) -> Path:
        """The file path of this package."""
        return self._path

    def _ensure_open(self) -> None:
        """Ensure the package is open."""
        if self._zf is None:
            raise OPCPackageError("Package is not open. Use OPCPackage.open() or context manager.")


class OPCPackageWriter:
    """Build a new OPC package from scratch.

    Usage::

        writer = OPCPackageWriter()
        writer.content_types.add_default("rels", CT_RELS)
        writer.content_types.add_default("xml", "application/xml")
        writer.content_types.add_override("/word/document.xml", CT_DOCUMENT)

        writer.root_rels.add(RT_OFFICE_DOCUMENT, "word/document.xml")

        writer.add_part("word/document.xml", document_xml_bytes)
        writer.add_rels("word/document.xml", doc_rels)

        writer.save("output.docx")
    """

    def __init__(self) -> None:
        self._content_types = ContentTypeMap()
        self._root_rels = RelationshipManager()
        self._parts: dict[str, bytes] = {}
        self._rels: dict[str, RelationshipManager] = {}

    @property
    def content_types(self) -> ContentTypeMap:
        """The content type map (mutable)."""
        return self._content_types

    @property
    def root_rels(self) -> RelationshipManager:
        """Root-level relationship manager (mutable)."""
        return self._root_rels

    def add_part(self, part_name: str, data: bytes) -> None:
        """Add or replace a part in the package."""
        normalized = part_name.lstrip("/")
        self._parts[normalized] = data

    def add_xml_part(self, part_name: str, root: etree._Element) -> None:
        """Add an XML part from an lxml element tree."""
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.add_part(part_name, data)

    def add_rels(self, source: str, rels: RelationshipManager) -> None:
        """Set the relationship manager for a source part."""
        self._rels[source] = rels

    def get_rels(self, source: str) -> RelationshipManager:
        """Get or create a relationship manager for a source part."""
        if source not in self._rels:
            self._rels[source] = RelationshipManager()
        return self._rels[source]

    def save(self, path: str | Path) -> Path:
        """Write the package to a ZIP file.

        Returns:
            The output path.
        """
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        data = self.save_bytes()
        path.write_bytes(data)
        return path

    def save_bytes(self) -> bytes:
        """Write the package to bytes (in-memory ZIP)."""
        from io import BytesIO

        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # Content types
            zf.writestr("[Content_Types].xml", self._content_types.serialize())

            # Root rels
            zf.writestr("_rels/.rels", self._root_rels.serialize())

            # Parts
            for name, data in sorted(self._parts.items()):
                zf.writestr(name, data)

            # Part-level rels
            for source, mgr in sorted(self._rels.items()):
                rels_path = _rels_path_for(source)
                zf.writestr(rels_path, mgr.serialize())

        return buf.getvalue()


def _rels_path_for(source: str) -> str:
    """Compute the .rels file path for a given source part.

    Root ("/") → "_rels/.rels"
    "word/document.xml" → "word/_rels/document.xml.rels"
    """
    if source == "/":
        return "_rels/.rels"

    source = source.lstrip("/")
    p = PurePosixPath(source)
    return str(p.parent / "_rels" / (p.name + ".rels"))
