from pathlib import Path

import pytest

from parsers.xls_parser import XlsParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_xls_parse_fixture() -> None:
    parser = XlsParser()
    doc = parser.parse(str(FIXTURES / "sample_a.xls"))
    ids = [seg.id for seg in doc.segments]
    assert "SheetA!A1" in ids
    assert "SheetA!B2" in ids
    assert "SheetB!C3" in ids
    assert "SheetB!D4" in ids


def test_xls_can_handle() -> None:
    parser = XlsParser()
    assert parser.can_handle("file.xls") is True
    assert parser.can_handle("file.xlsx") is False


def test_xls_rejects_invalid_source_column() -> None:
    parser = XlsParser()
    with pytest.raises(ValueError):
        parser.set_source_column("A-1")
