from __future__ import annotations

from pathlib import Path

import pytest

from core.models import ParseError, UnsupportedFormatError
from core.orchestrator import Orchestrator


FIXTURES = Path(__file__).resolve().parent / "fixtures"


def test_orchestrator_xliff(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(
        str(FIXTURES / "sample_a.xliff"),
        str(FIXTURES / "sample_b.xliff"),
        str(tmp_path),
    )
    assert len(outputs) == 2
    html_file = Path(outputs[0])
    excel_file = Path(outputs[1])
    assert html_file.exists()
    assert excel_file.exists()
    assert html_file.suffix == ".html"
    assert excel_file.suffix == ".xlsx"


def test_orchestrator_txt(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(
        str(FIXTURES / "sample_a.txt"),
        str(FIXTURES / "sample_b.txt"),
        str(tmp_path),
    )
    assert len(outputs) == 2
    html_file = Path(outputs[0])
    excel_file = Path(outputs[1])
    assert html_file.exists()
    assert excel_file.exists()
    assert html_file.suffix == ".html"
    assert excel_file.suffix == ".xlsx"


def test_orchestrator_unsupported_format(tmp_path: Path) -> None:
    file_a = tmp_path / "file_a.unknown"
    file_b = tmp_path / "file_b.unknown"
    file_a.write_text("a", encoding="utf-8")
    file_b.write_text("b", encoding="utf-8")
    orchestrator = Orchestrator()
    with pytest.raises(UnsupportedFormatError):
        orchestrator.compare_files(str(file_a), str(file_b), str(tmp_path))


def test_orchestrator_missing_file(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    missing = tmp_path / "missing.txt"
    other = tmp_path / "other.txt"
    other.write_text("content", encoding="utf-8")
    with pytest.raises(ParseError):
        orchestrator.compare_files(str(missing), str(other), str(tmp_path))
