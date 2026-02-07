from __future__ import annotations

from pathlib import Path
import shutil

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


def test_orchestrator_compare_folders(tmp_path: Path) -> None:
    folder_a = tmp_path / "a"
    folder_b = tmp_path / "b"
    folder_a.mkdir()
    folder_b.mkdir()

    shutil.copy(FIXTURES / "sample_a.txt", folder_a / "shared.txt")
    shutil.copy(FIXTURES / "sample_b.txt", folder_b / "shared.txt")

    shutil.copy(FIXTURES / "sample_a.xliff", folder_a / "only_a.xliff")
    shutil.copy(FIXTURES / "sample_b.srt", folder_b / "only_b.srt")

    (folder_a / "bad.unknown").write_text("a", encoding="utf-8")
    (folder_b / "bad.unknown").write_text("b", encoding="utf-8")

    orchestrator = Orchestrator()
    batch = orchestrator.compare_folders(str(folder_a), str(folder_b), str(tmp_path))

    assert batch.total_files == 4
    assert batch.compared_files == 1
    assert batch.only_in_a == 1
    assert batch.only_in_b == 1
    assert batch.errors == 1
    assert batch.summary_report_path is not None
    assert Path(batch.summary_report_path).exists()


def test_orchestrator_compare_folders_empty(tmp_path: Path) -> None:
    folder_a = tmp_path / "a"
    folder_b = tmp_path / "b"
    folder_a.mkdir()
    folder_b.mkdir()
    orchestrator = Orchestrator()
    batch = orchestrator.compare_folders(str(folder_a), str(folder_b), str(tmp_path))
    assert batch.total_files == 0
    assert Path(batch.summary_report_path).exists()


def test_orchestrator_compare_versions(tmp_path: Path) -> None:
    v1 = tmp_path / "v1.txt"
    v2 = tmp_path / "v2.txt"
    v3 = tmp_path / "v3.txt"
    v1.write_text("line one\n", encoding="utf-8")
    v2.write_text("line one changed\n", encoding="utf-8")
    v3.write_text("line one changed again\n", encoding="utf-8")

    orchestrator = Orchestrator()
    result = orchestrator.compare_versions([str(v1), str(v2), str(v3)], str(tmp_path))
    assert len(result.comparisons) == 2
    assert result.summary_report_path is not None
    assert Path(result.summary_report_path).exists()
