from pathlib import Path

import pytest

from core.models import ParseError
from parsers.xliff_parser import XliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_xliff_parse_fixture_a() -> None:
    parser = XliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.xliff"))
    assert len(doc.segments) == 4
    assert [seg.id for seg in doc.segments] == ["1", "2", "3", "4"]
    assert doc.segments[0].source == "Hello"
    assert doc.segments[0].target == "Привет"


def test_xliff_parse_fixture_b() -> None:
    parser = XliffParser()
    doc = parser.parse(str(FIXTURES / "sample_b.xliff"))
    assert len(doc.segments) == 4
    ids = [seg.id for seg in doc.segments]
    assert "5" in ids
    target_map = {seg.id: seg.target for seg in doc.segments}
    assert target_map["2"] == "Мир!"


def test_xliff_can_handle() -> None:
    parser = XliffParser()
    assert parser.can_handle("file.xliff") is True
    assert parser.can_handle("file.xlf") is True
    assert parser.can_handle("file.txt") is False


def test_xliff_empty_file(tmp_path: Path) -> None:
    empty_file = tmp_path / "empty.xliff"
    empty_file.write_text("", encoding="utf-8")
    parser = XliffParser()
    with pytest.raises(ParseError):
        parser.parse(str(empty_file))
