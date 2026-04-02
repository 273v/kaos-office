"""DOCX Metadata extraction.

Parses docProps/core.xml (Dublin Core) and docProps/app.xml (Application properties)
into DocumentMetadata.
"""

from __future__ import annotations

from dataclasses import dataclass

from kaos_office.ooxml.namespace import DC, DCTERMS, EP, qn
from kaos_office.opc.security import parse_xml_safe


@dataclass(frozen=True, slots=True)
class DocxMetadata:
    """Metadata extracted from DOCX core and app properties."""

    title: str | None = None
    creator: str | None = None
    description: str | None = None
    subject: str | None = None
    created: str | None = None
    modified: str | None = None
    last_modified_by: str | None = None
    revision: str | None = None
    # App properties
    word_count: int | None = None
    page_count: int | None = None
    paragraph_count: int | None = None
    company: str | None = None
    application: str | None = None

    @classmethod
    def from_xml(
        cls,
        core_xml: bytes | None = None,
        app_xml: bytes | None = None,
    ) -> DocxMetadata:
        """Create metadata from core.xml and app.xml bytes.

        Args:
            core_xml: Raw bytes of docProps/core.xml, or None.
            app_xml: Raw bytes of docProps/app.xml, or None.

        Returns:
            DocxMetadata with available fields populated.
        """
        title = None
        creator = None
        description = None
        subject = None
        created = None
        modified = None
        last_modified_by = None
        revision = None

        if core_xml is not None:
            root = parse_xml_safe(core_xml)
            title = _text(root, qn(DC, "title"))
            creator = _text(root, qn(DC, "creator"))
            description = _text(root, qn(DC, "description"))
            subject = _text(root, qn(DC, "subject"))
            created = _text(root, qn(DCTERMS, "created"))
            modified = _text(root, qn(DCTERMS, "modified"))
            # lastModifiedBy uses the cp namespace
            last_modified_by = _text(
                root,
                qn(
                    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                    "lastModifiedBy",
                ),
            )
            revision = _text(
                root,
                qn(
                    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                    "revision",
                ),
            )

        word_count = None
        page_count = None
        paragraph_count = None
        company = None
        application = None

        if app_xml is not None:
            root = parse_xml_safe(app_xml)
            word_count = _int_text(root, qn(EP, "Words"))
            page_count = _int_text(root, qn(EP, "Pages"))
            paragraph_count = _int_text(root, qn(EP, "Paragraphs"))
            company = _text(root, qn(EP, "Company"))
            application = _text(root, qn(EP, "Application"))

        return cls(
            title=title,
            creator=creator,
            description=description,
            subject=subject,
            created=created,
            modified=modified,
            last_modified_by=last_modified_by,
            revision=revision,
            word_count=word_count,
            page_count=page_count,
            paragraph_count=paragraph_count,
            company=company,
            application=application,
        )

    def to_dict(self) -> dict[str, object]:
        """Convert to a dict, omitting None values."""
        import dataclasses

        return {
            f.name: getattr(self, f.name)
            for f in dataclasses.fields(self)
            if getattr(self, f.name) is not None
        }


def _text(root: object, tag: str) -> str | None:
    """Find an element and return its text, or None."""
    from lxml import etree

    assert isinstance(root, etree._Element)
    el = root.find(f".//{tag}")
    if el is not None and el.text:
        return el.text.strip()
    return None


def _int_text(root: object, tag: str) -> int | None:
    """Find an element and return its text as int, or None."""
    val = _text(root, tag)
    if val is not None:
        try:
            return int(val)
        except ValueError:
            return None
    return None
