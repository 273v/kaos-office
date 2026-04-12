"""Office document error hierarchy for kaos-office.

All errors subclass KaosOfficeError → KaosCoreError, carrying structured
details for agent-friendly error messages and middleware decision-making.
"""

from __future__ import annotations

from kaos_core.exceptions import KaosCoreError


class KaosOfficeError(KaosCoreError):
    """Base error for all kaos-office operations."""


class DocxExtractionError(KaosOfficeError):
    """DOCX extraction failed (corrupt file, unsupported feature, etc.)."""


class PptxExtractionError(KaosOfficeError):
    """PPTX extraction failed (corrupt file, unsupported shape type, etc.)."""


class XlsxExtractionError(KaosOfficeError):
    """XLSX extraction failed (corrupt file, unsupported format, etc.)."""
