from pathlib import Path

from parsers.sdlxliff_parser import SdlXliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_sdlxliff_parse_fixture() -> None:
    parser = SdlXliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.sdlxliff"))
    assert len(doc.segments) == 3
    assert doc.segments[0].id == "101"
    assert doc.segments[0].source == "First line"
    assert doc.segments[0].target == "Первая строка"
    assert doc.segments[0].metadata.get("conf") == "ApprovedTranslation"


def test_sdlxliff_can_handle() -> None:
    parser = SdlXliffParser()
    assert parser.can_handle("file.sdlxliff") is True
    assert parser.can_handle("file.xliff") is False
