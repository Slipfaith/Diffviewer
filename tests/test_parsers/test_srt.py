from pathlib import Path

from parsers.srt_parser import SrtParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_srt_parse_fixture() -> None:
    parser = SrtParser()
    doc = parser.parse(str(FIXTURES / "sample_a.srt"))
    assert len(doc.segments) == 4
    assert doc.segments[0].id == "1"
    assert doc.segments[1].target == "Second subtitle line\nthat can span multiple lines"
    assert doc.segments[1].metadata["start"] == "00:00:05,000"
    assert doc.segments[1].metadata["end"] == "00:00:08,000"


def test_srt_can_handle() -> None:
    parser = SrtParser()
    assert parser.can_handle("file.srt") is True
    assert parser.can_handle("file.txt") is False
