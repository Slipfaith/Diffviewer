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
        assert segment.id == f"para_{i}", f"Expected para_{i}, got {segment.id}"


def test_doc_validate_returns_error() -> None:
    parser = DocxParser()
    errors = parser.validate("sample.doc")
    assert any("DOC format" in e for e in errors)
