from pathlib import Path

import pytest

from parsers.docx_parser import DocxParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_docx_parse_fixture() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_a.docx"))
    assert len(doc.segments) == 3
    assert doc.segments[0].id == "body_p1"
    assert "style" in doc.segments[0].metadata


def test_docx_preserves_whitespace_tokens() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_whitespace.docx"))
    assert any(segment.target == "Alpha  Beta\tGamma" for segment in doc.segments)
    assert any(segment.target.endswith("  ") for segment in doc.segments)


def test_docx_can_handle() -> None:
    parser = DocxParser()
    assert parser.can_handle("file.docx") is True
    assert parser.can_handle("file.doc") is False


def test_doc_not_in_supported_extensions() -> None:
    parser = DocxParser()
    assert ".doc" not in parser.supported_extensions
    assert ".docx" in parser.supported_extensions


def test_docx_sequential_segment_ids() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_a.docx"))
    for i, segment in enumerate(doc.segments, start=1):
        assert segment.id == f"body_p{i}", f"Expected body_p{i}, got {segment.id}"


def test_doc_validate_returns_error() -> None:
    parser = DocxParser()
    errors = parser.validate("sample.doc")
    assert any("DOC format" in e for e in errors)


def test_docx_table_cells_extracted() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_full.docx"))
    ids = [s.id for s in doc.segments]
    # Table cells must be present
    assert "t1_r1_c1_p1" in ids
    assert "t1_r1_c2_p1" in ids
    assert "t1_r2_c1_p1" in ids
    assert "t1_r2_c2_p1" in ids
    texts = [s.target for s in doc.segments]
    assert "Cell A1" in texts
    assert "Cell B2" in texts


def test_docx_table_preserves_body_order() -> None:
    """Paragraph after table must not appear before table cells."""
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_full.docx"))
    ids = [s.id for s in doc.segments]
    assert ids.index("t1_r1_c1_p1") < ids.index("body_p3")


def test_docx_header_footer_extracted() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_full.docx"))
    ids = [s.id for s in doc.segments]
    assert "hdr_s1_p1" in ids
    assert "ftr_s1_p1" in ids
    texts = [s.target for s in doc.segments]
    assert "Header text" in texts
    assert "Footer text" in texts


def test_docx_textbox_extracted_separately() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_textbox.docx"))
    ids = [s.id for s in doc.segments]
    assert "txbx1_p1" in ids
    texts = [s.target for s in doc.segments]
    assert "Text box content here" in texts


def test_docx_textbox_not_mixed_into_anchor_paragraph() -> None:
    """Text box text must not appear inside the paragraph that anchors the shape."""
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_textbox.docx"))
    body_segs = [s for s in doc.segments if s.id.startswith("body_")]
    for seg in body_segs:
        assert "Text box content here" not in seg.target


def test_docx_footnote_extracted() -> None:
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_footnote.docx"))
    ids = [s.id for s in doc.segments]
    assert "fn1_p1" in ids
    texts = [s.target for s in doc.segments]
    assert "This is footnote number one" in texts


def test_docx_footnote_separators_skipped() -> None:
    """Separator notes (id=-1, id=0) must not produce segments."""
    parser = DocxParser()
    doc = parser.parse(str(FIXTURES / "sample_footnote.docx"))
    ids = [s.id for s in doc.segments]
    assert not any("fn-1" in i or "fn0" in i for i in ids)
