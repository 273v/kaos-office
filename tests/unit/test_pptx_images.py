"""PPTX embedded-image extraction — bytes inlined as data-URIs.

Regression coverage for the fix that makes embedded PPTX pictures carry their
actual bytes (a ``data:image/<fmt>;base64,...`` URI by default) instead of a
bare ``pptx://name.ext`` placeholder, bringing PPTX to parity with the DOCX
reader so downstream OCR / VLM / writers can reach the image content.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from kaos_office.pptx.reader import parse_pptx

_IMAGES_PPTX = Path(__file__).parent.parent / "fixtures" / "pptx" / "battle" / "images.pptx"


def _image_srcs(doc: object) -> list[str | None]:
    """Collect every ``Image`` node's ``src`` from a ContentDocument."""
    srcs: list[str | None] = []
    seen: set[int] = set()

    def walk(node: object, depth: int = 0) -> None:
        if id(node) in seen or depth > 40:
            return
        seen.add(id(node))
        if type(node).__name__ == "Image":
            srcs.append(getattr(node, "src", None))
        for attr in vars(node) if hasattr(node, "__dict__") else []:
            value = getattr(node, attr, None)
            if isinstance(value, (list, tuple)):
                for item in value:
                    if hasattr(item, "__dict__"):
                        walk(item, depth + 1)
            elif hasattr(value, "__dict__"):
                walk(value, depth + 1)

    walk(doc)
    return srcs


@pytest.fixture
def images_pptx() -> Path:
    if not _IMAGES_PPTX.exists():
        pytest.skip(f"fixture not found: {_IMAGES_PPTX}")
    return _IMAGES_PPTX


class TestPptxImageBytes:
    def test_default_inlines_data_uris(self, images_pptx: Path) -> None:
        srcs = _image_srcs(parse_pptx(images_pptx))
        assert srcs, "fixture should contain embedded pictures"
        # Every embedded picture now carries its bytes inline; none is a bare
        # pptx:// placeholder.
        assert all(s and s.startswith("data:image/") for s in srcs), srcs
        assert not any(s and s.startswith("pptx://") for s in srcs)

    def test_mime_subtype_preserved(self, images_pptx: Path) -> None:
        # images.pptx carries both a PNG and a JPEG; the data-URI MIME subtype
        # must reflect each picture's real content type.
        srcs = [s for s in _image_srcs(parse_pptx(images_pptx)) if s]
        assert any(s.startswith("data:image/png;base64,") for s in srcs)
        assert any(s.startswith("data:image/jpeg;base64,") for s in srcs)

    def test_custom_image_src_builder(self, images_pptx: Path) -> None:
        # A caller-supplied builder receives (bytes, fmt, index) and its return
        # value becomes Image.src — the out-of-band-storage contract.
        seen: list[tuple[int, str, int]] = []

        def builder(data: bytes, fmt: str, index: int) -> str:
            seen.append((len(data), fmt, index))
            return f"vfs://img-{index}.{fmt}"

        srcs = [s for s in _image_srcs(parse_pptx(images_pptx, image_src_builder=builder)) if s]
        assert srcs == ["vfs://img-1.png", "vfs://img-2.jpeg"]
        # Builder saw real bytes and a 1-based index per picture.
        assert [fmt for _len, fmt, _i in seen] == ["png", "jpeg"]
        assert [i for _len, _fmt, i in seen] == [1, 2]
        assert all(length > 0 for length, _fmt, _i in seen)
