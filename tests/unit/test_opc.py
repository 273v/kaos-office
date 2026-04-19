"""Tests for OPC package layer: security, content types, relationships, package."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from kaos_office.opc.content_types import ContentTypeMap
from kaos_office.opc.package import OPCPackage, OPCPackageError
from kaos_office.opc.relationships import RelationshipManager
from kaos_office.opc.security import OPCSecurityError, _validate_part_path, validate_zip_security

# ──────────────────────────── Security ────────────────────────────


class TestPathValidation:
    def test_normal_path(self):
        _validate_part_path("word/document.xml")

    def test_rels_path(self):
        _validate_part_path("_rels/.rels")

    def test_absolute_path_rejected(self):
        with pytest.raises(OPCSecurityError, match="Absolute path"):
            _validate_part_path("/etc/passwd")

    def test_drive_letter_rejected(self):
        with pytest.raises(OPCSecurityError, match="Drive letter"):
            _validate_part_path("C:\\Windows\\system32")

    def test_traversal_rejected(self):
        with pytest.raises(OPCSecurityError, match="Path traversal"):
            _validate_part_path("word/../../etc/passwd")

    def test_hidden_file_rejected(self):
        with pytest.raises(OPCSecurityError, match="Hidden file"):
            _validate_part_path(".hidden/secret.xml")

    def test_rels_hidden_allowed(self):
        # .rels files are standard OPC
        _validate_part_path("word/_rels/.rels")


class TestZipSecurity:
    def _make_zip(self, entries: dict[str, bytes]) -> tuple[zipfile.ZipFile, int]:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            for name, data in entries.items():
                zf.writestr(name, data)
        buf.seek(0)
        size = len(buf.getvalue())
        zf = zipfile.ZipFile(buf, "r")
        return zf, size

    def test_normal_zip_passes(self):
        zf, size = self._make_zip({"test.xml": b"<root/>"})
        validate_zip_security(zf, size)
        zf.close()

    def test_traversal_in_zip(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            info = zipfile.ZipInfo("../../etc/passwd")
            zf.writestr(info, b"pwned")
        buf.seek(0)
        zf = zipfile.ZipFile(buf, "r")
        with pytest.raises(OPCSecurityError, match="Path traversal"):
            validate_zip_security(zf, len(buf.getvalue()))
        zf.close()


# ──────────────────────────── Content Types ────────────────────────────


class TestContentTypeMap:
    CT_XML = b"""\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    def test_parse(self):
        ct = ContentTypeMap.parse(self.CT_XML)
        assert ct.defaults["xml"] == "application/xml"
        assert ct.defaults["png"] == "image/png"

    def test_get_override(self):
        ct = ContentTypeMap.parse(self.CT_XML)
        ct_val = ct.get("word/document.xml")
        assert ct_val is not None
        assert "wordprocessingml" in ct_val

    def test_get_default_by_extension(self):
        ct = ContentTypeMap.parse(self.CT_XML)
        assert ct.get("word/media/image1.png") == "image/png"

    def test_get_unknown(self):
        ct = ContentTypeMap.parse(self.CT_XML)
        assert ct.get("word/unknown.xyz") is None

    def test_override_takes_precedence(self):
        ct = ContentTypeMap.parse(self.CT_XML)
        # document.xml has an override, so it should NOT return "application/xml"
        result = ct.get("word/document.xml")
        assert result != "application/xml"


# ──────────────────────────── Relationships ────────────────────────────


class TestRelationshipManager:
    RELS_XML = b"""\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com" TargetMode="External"/>
</Relationships>"""

    def test_parse(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        assert len(mgr) == 3

    def test_resolve(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        assert mgr.resolve("rId1") == "styles.xml"
        assert mgr.resolve("rId2") == "media/image1.png"

    def test_resolve_missing(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        assert mgr.resolve("rId99") is None

    def test_by_type(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        images = mgr.by_type(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        )
        assert len(images) == 1
        assert images[0].target == "media/image1.png"

    def test_first_target(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        target = mgr.first_target(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        )
        assert target == "styles.xml"

    def test_first_target_missing_type(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        assert mgr.first_target("http://nonexistent/type") is None

    def test_external_relationship(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        rel = mgr.get("rId3")
        assert rel is not None
        assert rel.external is True
        assert rel.target == "https://example.com"

    def test_contains(self):
        mgr = RelationshipManager.parse(self.RELS_XML)
        assert "rId1" in mgr
        assert "rId99" not in mgr

    def test_empty_rels(self):
        xml = b"""\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""
        mgr = RelationshipManager.parse(xml)
        assert len(mgr) == 0


# ──────────────────────────── OPCPackage ────────────────────────────


class TestOPCPackage:
    def _write_minimal_docx(self, tmp_path: Path) -> Path:
        from tests.conftest import make_minimal_docx

        docx_bytes = make_minimal_docx()
        path = tmp_path / "test.docx"
        path.write_bytes(docx_bytes)
        return path

    def test_open_and_close(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        pkg = OPCPackage.open(path)
        assert pkg.has_part("word/document.xml")
        pkg.close()

    def test_context_manager(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            assert pkg.has_part("word/document.xml")

    def test_read_part(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            data = pkg.read_part("word/document.xml")
            assert b"<w:document" in data

    def test_read_xml(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            root = pkg.read_xml("word/document.xml")
            assert root is not None

    def test_read_missing_part(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg, pytest.raises(OPCPackageError, match="Part not found"):
            pkg.read_part("word/nonexistent.xml")

    def test_has_part(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            assert pkg.has_part("word/document.xml")
            assert not pkg.has_part("word/nonexistent.xml")

    def test_list_parts(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            parts = pkg.list_parts()
            assert "[Content_Types].xml" in parts
            assert "word/document.xml" in parts

    def test_relationships_root(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            rels = pkg.relationships("/")
            target = rels.first_target(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            )
            assert target == "word/document.xml"

    def test_relationships_document(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            rels = pkg.relationships("word/document.xml")
            assert len(rels) >= 0  # May be empty if no styles etc.

    def test_content_type(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            ct = pkg.content_type("word/document.xml")
            assert ct is not None
            assert "wordprocessingml" in ct

    def test_file_not_found(self, tmp_path: Path):
        with pytest.raises(OPCPackageError, match="File not found"):
            OPCPackage.open(tmp_path / "nonexistent.docx")

    def test_invalid_zip(self, tmp_path: Path):
        path = tmp_path / "bad.docx"
        path.write_bytes(b"not a zip file")
        with pytest.raises(OPCPackageError, match="Not a valid ZIP"):
            OPCPackage.open(path)

    def test_missing_content_types(self, tmp_path: Path):
        path = tmp_path / "no_ct.docx"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("word/document.xml", b"<doc/>")
        path.write_bytes(buf.getvalue())
        with pytest.raises(OPCPackageError, match=r"Missing \[Content_Types\].xml"):
            OPCPackage.open(path)

    def test_parts_cached(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        with OPCPackage.open(path) as pkg:
            data1 = pkg.read_part("word/document.xml")
            data2 = pkg.read_part("word/document.xml")
            assert data1 is data2  # Same object from cache

    def test_closed_package_raises(self, tmp_path: Path):
        path = self._write_minimal_docx(tmp_path)
        pkg = OPCPackage.open(path)
        pkg.close()
        with pytest.raises(OPCPackageError, match="not open"):
            pkg.read_part("word/document.xml")
