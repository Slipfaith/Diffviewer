from pathlib import Path

from parsers.pptx_parser import PptxParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_pptx_parse_fixture() -> None:
    parser = PptxParser()
    doc = parser.parse(str(FIXTURES / "sample_a.pptx"))
    assert len(doc.segments) >= 2
    assert any(seg.id.startswith("slide1_") for seg in doc.segments)
    assert any(seg.id.startswith("slide2_") for seg in doc.segments)


def test_pptx_can_handle() -> None:
    parser = PptxParser()
    assert parser.can_handle("file.pptx") is True
    assert parser.can_handle("file.docx") is False
