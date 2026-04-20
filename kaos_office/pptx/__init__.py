"""PPTX (PresentationML) extraction and generation.

The writer depends on the optional ``[pptx]`` extra (``python-pptx``).
Importing this subpackage without that extra used to crash the entire
``kaos_office`` import path (parent ``__init__`` re-exports from here).
Guard the writer import so only ``write_pptx`` / ``write_pptx_bytes``
become unavailable when the extra is absent; every other kaos-office
entry point (``parse_docx``, XLSX reader/writer, etc.) remains usable.
"""

try:
    from kaos_office.pptx.writer import write_pptx, write_pptx_bytes
except ImportError:  # python-pptx not installed — writer unavailable
    write_pptx = None  # type: ignore[assignment]
    write_pptx_bytes = None  # type: ignore[assignment]

__all__ = ["write_pptx", "write_pptx_bytes"]
