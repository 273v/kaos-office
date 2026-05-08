"""PPTX (PresentationML) extraction and generation.

Reader: :func:`parse_pptx` (lazy-imports ``python-pptx``; ImportError
with the ``[pptx]`` install hint at call time) plus the cheap
inspectors :func:`get_slide_count`, :func:`get_slide_text`,
:func:`get_slide_notes`, and :func:`list_slides`.

Writer: :func:`write_pptx` and :func:`write_pptx_bytes` are typed
lazy wrappers around :mod:`kaos_office.pptx.writer`. The writer
depends on the optional ``[pptx]`` extra (``python-pptx``); importing
this subpackage without that extra used to crash the entire
``kaos_office`` import path. The wrappers defer the
``kaos_office.pptx.writer`` import (and therefore the ``python-pptx``
import) until call time and raise :class:`ImportError` with an
explicit install hint instead of materializing as ``None`` callables.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Literal

from kaos_office.pptx.reader import (
    get_slide_count,
    get_slide_notes,
    get_slide_text,
    list_slides,
    parse_pptx,
)

if TYPE_CHECKING:
    from kaos_content import ContentDocument

OverflowMode = Literal["warn", "autofit", "extend"]

_INSTALL_HINT = (
    "PPTX writing requires the `python-pptx` package. "
    "Install it with: uv add 'kaos-office[pptx]' (or pip install 'kaos-office[pptx]')."
)


def write_pptx(
    doc: ContentDocument,
    path: str | Path,
    *,
    template: str | Path | None = None,
    overflow: OverflowMode = "warn",
) -> Path:
    """Write a ContentDocument to a PPTX file.

    Lazy wrapper around :func:`kaos_office.pptx.writer.write_pptx` that
    raises :class:`ImportError` with the ``[pptx]`` install hint when
    ``python-pptx`` is not installed. See the underlying function for
    full ``template`` / ``overflow`` semantics.
    """
    try:
        from kaos_office.pptx.writer import write_pptx as _impl
    except ImportError as exc:  # pragma: no cover — covered by ImportError path test
        raise ImportError(_INSTALL_HINT) from exc
    return _impl(doc, path, template=template, overflow=overflow)


def write_pptx_bytes(
    doc: ContentDocument,
    *,
    template: str | Path | None = None,
    overflow: OverflowMode = "warn",
) -> bytes:
    """Write a ContentDocument to PPTX bytes (in-memory).

    Lazy wrapper around :func:`kaos_office.pptx.writer.write_pptx_bytes`
    that raises :class:`ImportError` with the ``[pptx]`` install hint
    when ``python-pptx`` is not installed.
    """
    try:
        from kaos_office.pptx.writer import write_pptx_bytes as _impl
    except ImportError as exc:  # pragma: no cover — covered by ImportError path test
        raise ImportError(_INSTALL_HINT) from exc
    return _impl(doc, template=template, overflow=overflow)


__all__ = [
    "OverflowMode",
    "get_slide_count",
    "get_slide_notes",
    "get_slide_text",
    "list_slides",
    "parse_pptx",
    "write_pptx",
    "write_pptx_bytes",
]
