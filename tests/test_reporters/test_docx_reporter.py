from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

from core.models import ChangeStatistics, ComparisonResult, ParsedDocument
from reporters.docx_reporter import DocxTrackChangesReporter


def make_result(tmp_path: Path) -> ComparisonResult:
    file_a = tmp_path / "a.docx"
    file_b = tmp_path / "b.docx"
    file_a.write_text("a", encoding="utf-8")
    file_b.write_text("b", encoding="utf-8")
    doc_a = ParsedDocument([], "DOCX", str(file_a))
    doc_b = ParsedDocument([], "DOCX", str(file_b))
    return ComparisonResult(
        file_a=doc_a,
        file_b=doc_b,
        changes=[],
        statistics=ChangeStatistics.from_changes([]),
        timestamp=datetime.now(timezone.utc),
    )


def test_docx_reporter_unavailable_fallback(tmp_path: Path) -> None:
    result = make_result(tmp_path)
    reporter = DocxTrackChangesReporter()
    reporter._exe_path = tmp_path / "missing.exe"
    assert reporter.is_available() is False

    output = reporter.generate(result, str(tmp_path / "report.docx"))
    assert Path(output).suffix == ".html"
    assert Path(output).exists()
    assert (tmp_path / "report.xlsx").exists()


def test_docx_reporter_subprocess_called(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    result = make_result(tmp_path)
    exe_path = tmp_path / "docx_compare.exe"
    exe_path.write_text("dummy", encoding="utf-8")
    reporter = DocxTrackChangesReporter(author="Tester")
    reporter._exe_path = exe_path

    called = {}

    def fake_run(cmd, capture_output=True, text=True):
        called["cmd"] = cmd
        return SimpleNamespace(returncode=0, stderr="", stdout="")

    monkeypatch.setattr("reporters.docx_reporter.subprocess.run", fake_run)
    output = reporter.generate(result, str(tmp_path / "out.docx"))
    assert Path(output).suffix == ".docx"
    assert called["cmd"][0] == str(exe_path)
    assert called["cmd"][1] == result.file_a.file_path
    assert called["cmd"][2] == result.file_b.file_path
    assert called["cmd"][3].endswith("out.docx")
    assert called["cmd"][4] == "--author"
    assert called["cmd"][5] == "Tester"


def test_docx_reporter_nonzero_exit(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    result = make_result(tmp_path)
    exe_path = tmp_path / "docx_compare.exe"
    exe_path.write_text("dummy", encoding="utf-8")
    reporter = DocxTrackChangesReporter()
    reporter._exe_path = exe_path

    def fake_run(cmd, capture_output=True, text=True):
        return SimpleNamespace(returncode=1, stderr="boom", stdout="")

    monkeypatch.setattr("reporters.docx_reporter.subprocess.run", fake_run)
    with pytest.raises(RuntimeError):
        reporter.generate(result, str(tmp_path / "out.docx"))
