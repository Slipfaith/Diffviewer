from pathlib import Path

from parsers.memoq_parser import MemoQXliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_memoq_parse_fixture() -> None:
    parser = MemoQXliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.mqxliff"))
    assert len(doc.segments) == 3
    seg = doc.segments[0]
    assert seg.id == "1"
    assert seg.source.startswith("Log in")
    assert seg.target.startswith("Log elke")
    assert seg.metadata.get("status") == "ManuallyConfirmed"
    assert seg.metadata.get("segmentguid") == "guid-1"
    assert seg.metadata.get("context") == ":B5/C5"


def test_memoq_can_handle() -> None:
    parser = MemoQXliffParser()
    assert parser.can_handle("file.mqxliff") is True
    assert parser.can_handle("file.xliff") is False
