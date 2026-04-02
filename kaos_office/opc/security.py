"""OPC package security checks.

Protects against ZIP bombs, path traversal, XML bombs, and oversized files.
"""

from __future__ import annotations

import zipfile
from pathlib import PurePosixPath

from lxml import etree

# --- Size limits ---
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB compressed
MAX_UNCOMPRESSED_TOTAL = 500 * 1024 * 1024  # 500 MB total uncompressed
MAX_PART_SIZE = 50 * 1024 * 1024  # 50 MB single part
MAX_COMPRESSION_RATIO = 100  # Reject if ratio > 100:1


class OPCSecurityError(Exception):
    """Raised when an OPC package fails security validation."""


def validate_zip_security(zf: zipfile.ZipFile, file_size: int) -> None:
    """Validate a ZIP file against security constraints.

    Args:
        zf: Open ZipFile to validate.
        file_size: Size of the compressed file on disk.

    Raises:
        OPCSecurityError: If any security check fails.
    """
    if file_size > MAX_FILE_SIZE:
        raise OPCSecurityError(
            f"File size {file_size:,} bytes exceeds maximum {MAX_FILE_SIZE:,} bytes. "
            "If this is a legitimate file, increase MAX_FILE_SIZE."
        )

    total_uncompressed = 0
    for info in zf.infolist():
        # Path traversal check
        _validate_part_path(info.filename)

        # Individual part size check
        if info.file_size > MAX_PART_SIZE:
            raise OPCSecurityError(
                f"Part '{info.filename}' uncompressed size {info.file_size:,} bytes "
                f"exceeds maximum {MAX_PART_SIZE:,} bytes."
            )

        # Compression ratio check (ZIP bomb detection)
        if info.compress_size > 0:
            ratio = info.file_size / info.compress_size
            if ratio > MAX_COMPRESSION_RATIO:
                raise OPCSecurityError(
                    f"Part '{info.filename}' has suspicious compression ratio {ratio:.0f}:1 "
                    f"(max {MAX_COMPRESSION_RATIO}:1). Possible ZIP bomb."
                )

        total_uncompressed += info.file_size

    # Total uncompressed size check
    if total_uncompressed > MAX_UNCOMPRESSED_TOTAL:
        raise OPCSecurityError(
            f"Total uncompressed size {total_uncompressed:,} bytes "
            f"exceeds maximum {MAX_UNCOMPRESSED_TOTAL:,} bytes."
        )


def _validate_part_path(path: str) -> None:
    """Reject paths that could escape the archive.

    Raises:
        OPCSecurityError: If path is suspicious.
    """
    # Absolute paths
    if path.startswith("/") or path.startswith("\\"):
        raise OPCSecurityError(f"Absolute path in archive: '{path}'")

    # Drive letters (Windows)
    if len(path) >= 2 and path[1] == ":":
        raise OPCSecurityError(f"Drive letter in archive path: '{path}'")

    # Parent directory traversal
    parts = PurePosixPath(path).parts
    if ".." in parts:
        raise OPCSecurityError(f"Path traversal in archive: '{path}'")

    # Hidden files (not standard in OPC, suspicious)
    for part in parts:
        if part.startswith(".") and part not in (".", "..") and not part.endswith(".rels"):
            raise OPCSecurityError(f"Hidden file in archive: '{path}'")


def safe_xml_parser() -> etree.XMLParser:
    """Create a secure XML parser that prevents XML bomb attacks.

    Returns:
        An lxml XMLParser configured to prevent entity expansion attacks.
    """
    return etree.XMLParser(
        resolve_entities=False,
        huge_tree=False,
        no_network=True,
        remove_comments=True,
    )


def parse_xml_safe(xml_bytes: bytes) -> etree._Element:
    """Parse XML bytes using a secure parser.

    Args:
        xml_bytes: Raw XML bytes.

    Returns:
        Parsed XML root element.

    Raises:
        etree.XMLSyntaxError: If XML is malformed.
    """
    return etree.fromstring(xml_bytes, parser=safe_xml_parser())
