from pathlib import Path

import pytest

from parsers.docx_parser import DocxParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_docx_parse_fixture() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_a.docx"))
    assert len(doc.segments) == 3
    assert doc.segments[0].id == "para_1"
    assert "style" in doc.segments[0].metadata


def test_docx_can_handle() -> None:
    parser = DocxParser()
    assert parser.can_handle("file.docx") is True
    assert parser.can_handle("file.doc") is True


def test_doc_parser_doc_not_implemented(tmp_path: Path) -> None:
    fake_doc = tmp_path / "sample.doc"
    fake_doc.write_text("dummy", encoding="utf-8")
    parser = DocxParser()
    with pytest.raises(NotImplementedError):
        parser.parse(str(fake_doc))
