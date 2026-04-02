"""Open Packaging Conventions (OPC) layer — shared by DOCX, XLSX, PPTX."""

from kaos_office.opc.content_types import ContentTypeMap
from kaos_office.opc.package import OPCPackage
from kaos_office.opc.relationships import Relationship, RelationshipManager

__all__ = [
    "ContentTypeMap",
    "OPCPackage",
    "Relationship",
    "RelationshipManager",
]
