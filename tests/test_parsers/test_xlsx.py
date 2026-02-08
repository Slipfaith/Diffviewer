from pathlib import Path

import openpyxl
import pytest

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


def test_xlsx_parse_with_source_column(tmp_path: Path) -> None:
    file_path = tmp_path / "bilingual.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "Source 1"
    sheet["B1"] = "Target 1"
    sheet["C1"] = "Alt target 1"
    sheet["A2"] = "Source 2"
    sheet["B2"] = "Target 2"
    workbook.save(file_path)

    parser = XlsxParser()
    parser.set_source_column("A")
    doc = parser.parse(str(file_path))
    by_id = {seg.id: seg for seg in doc.segments}

    assert "Sheet1!A1" not in by_id
    assert by_id["Sheet1!B1"].source == "Source 1"
    assert by_id["Sheet1!C1"].source == "Source 1"
    assert by_id["Sheet1!B2"].source == "Source 2"


def test_xlsx_rejects_invalid_source_column() -> None:
    parser = XlsxParser()
    with pytest.raises(ValueError):
        parser.set_source_column("A-1")
