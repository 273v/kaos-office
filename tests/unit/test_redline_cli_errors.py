"""Error and flag paths for the ``kaos-office redline`` CLI command."""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.revision import Revisions

from kaos_office.cli import main
from kaos_office.docx import parse_docx, write_docx


def _make_docx(path: Path, *paras: str) -> Path:
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=tuple(Paragraph(children=(Text(value=p),)) for p in paras),
    )
    write_docx(doc, path)
    return path


@pytest.fixture
def pair(tmp_path: Path) -> tuple[Path, Path]:
    a = _make_docx(tmp_path / "a.docx", "one", "two")
    b = _make_docx(tmp_path / "b.docx", "one", "two", "three")
    return a, b


def test_missing_original_exits_nonzero(pair: tuple[Path, Path], tmp_path: Path) -> None:
    _a, b = pair
    with pytest.raises(SystemExit) as exc:
        main(["redline", str(tmp_path / "missing.docx"), str(b), str(tmp_path / "out.docx")])
    assert exc.value.code != 0
    assert not (tmp_path / "out.docx").exists()


def test_non_docx_input_exits_nonzero(pair: tuple[Path, Path], tmp_path: Path) -> None:
    a, _b = pair
    bad = tmp_path / "bad.txt"
    bad.write_text("not a docx", encoding="utf-8")
    with pytest.raises(SystemExit) as exc:
        main(["redline", str(a), str(bad), str(tmp_path / "out.docx")])
    assert exc.value.code != 0


def test_no_moves_flag_runs_and_writes(pair: tuple[Path, Path], tmp_path: Path) -> None:
    a, b = pair
    out = tmp_path / "nm.docx"
    main(["redline", str(a), str(b), str(out), "--no-moves"])
    assert out.exists()
    assert Revisions.from_document(parse_docx(out, track_changes=True))
