from pathlib import Path

from parsers.xlsx_parser import XlsxParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_xlsx_parse_fixture() -> None:
    parser = XlsxParser()
    doc = parser.parse(str(FIXTURES / "sample_a.xlsx"))
    ids = [seg.id for seg in doc.segments]
    assert "SheetA!A1" in ids
    assert "SheetA!B2" in ids
    assert "SheetB!C3" in ids
    assert "SheetB!D4" in ids


def test_xlsx_can_handle() -> None:
    parser = XlsxParser()
    assert parser.can_handle("file.xlsx") is True
    assert parser.can_handle("file.xls") is False
