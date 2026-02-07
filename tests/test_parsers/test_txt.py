from pathlib import Path

from parsers.txt_parser import TxtParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_txt_parse_fixture_a() -> None:
    parser = TxtParser()
    doc = parser.parse(str(FIXTURES / "sample_a.txt"))
    assert len(doc.segments) == 6
    assert doc.segments[0].id == "1"
    assert doc.segments[2].target == ""
    assert doc.segments[3].target == "Привет мир"


def test_txt_parse_fixture_b() -> None:
    parser = TxtParser()
    doc = parser.parse(str(FIXTURES / "sample_b.txt"))
    assert len(doc.segments) == 6
    assert doc.segments[1].target == "Line two changed"


def test_txt_can_handle() -> None:
    parser = TxtParser()
    assert parser.can_handle("file.txt") is True
    assert parser.can_handle("file.xliff") is False


def test_txt_empty_file(tmp_path: Path) -> None:
    empty_file = tmp_path / "empty.txt"
    empty_file.write_text("", encoding="utf-8")
    parser = TxtParser()
    doc = parser.parse(str(empty_file))
    assert doc.segments == []


def test_txt_utf8_cyrillic(tmp_path: Path) -> None:
    path = tmp_path / "cyrillic.txt"
    path.write_text("Привет\nмир", encoding="utf-8")
    parser = TxtParser()
    doc = parser.parse(str(path))
    assert doc.segments[0].target == "Привет"
    assert doc.segments[1].target == "мир"
