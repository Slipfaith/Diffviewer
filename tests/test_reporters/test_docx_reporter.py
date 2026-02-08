from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

from core.models import ChangeStatistics, ComparisonResult, ParsedDocument
from reporters import docx_reporter
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


class FakeDoc:
    def __init__(self) -> None:
        self.closed = False
        self.saved_args = None

    def Close(self, SaveChanges=0) -> None:
        self.closed = True

    def SaveAs2(self, path, FileFormat=None) -> None:
        self.saved_args = (path, FileFormat)


class FakeDocuments:
    def __init__(self) -> None:
        self.opened = []

    def Open(self, path, ReadOnly=True):
        self.opened.append((path, ReadOnly))
        return FakeDoc()


class FakeWord:
    def __init__(self) -> None:
        self.Visible = None
        self.DisplayAlerts = None
        self.Documents = FakeDocuments()
        self.Application = self
        self.ActiveDocument = None
        self.compare_kwargs = None
        self.quit_called = False
        self.Version = "16.0"

    def CompareDocuments(self, **kwargs) -> None:
        self.compare_kwargs = kwargs
        self.ActiveDocument = FakeDoc()

    def Quit(self) -> None:
        self.quit_called = True


def test_docx_reporter_compare_documents_called(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    result = make_result(tmp_path)
    fake_word = FakeWord()

    def fake_dispatch(_):
        return fake_word

    monkeypatch.setattr(
        docx_reporter, "win32com", SimpleNamespace(client=SimpleNamespace(Dispatch=fake_dispatch))
    )
    monkeypatch.setattr(
        docx_reporter, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    )

    reporter = DocxTrackChangesReporter(author="Tester")
    output = reporter.generate(result, str(tmp_path / "out.docx"))
    assert Path(output).suffix == ".docx"
    assert fake_word.compare_kwargs is not None
    assert fake_word.compare_kwargs["Destination"] == 2
    assert fake_word.compare_kwargs["Granularity"] == 1
    assert fake_word.compare_kwargs["CompareFormatting"] is True
    assert fake_word.compare_kwargs["CompareCaseChanges"] is True
    assert fake_word.compare_kwargs["CompareWhitespace"] is True
    assert fake_word.compare_kwargs["CompareTables"] is True
    assert fake_word.compare_kwargs["CompareHeaders"] is True
    assert fake_word.compare_kwargs["CompareFootnotes"] is True
    assert fake_word.compare_kwargs["CompareTextboxes"] is True
    assert fake_word.compare_kwargs["CompareFields"] is True
    assert fake_word.compare_kwargs["CompareComments"] is True
    assert fake_word.compare_kwargs["RevisedAuthor"] == "Tester"
    assert fake_word.ActiveDocument is not None
    assert fake_word.ActiveDocument.saved_args is not None
    assert fake_word.ActiveDocument.saved_args[1] == 12
    assert fake_word.quit_called is True


def test_docx_reporter_falls_back_on_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    result = make_result(tmp_path)
    fake_word = FakeWord()

    def fake_dispatch(_):
        return fake_word

    def raise_compare(**_kwargs):
        raise RuntimeError("boom")

    fake_word.CompareDocuments = raise_compare

    monkeypatch.setattr(
        docx_reporter, "win32com", SimpleNamespace(client=SimpleNamespace(Dispatch=fake_dispatch))
    )
    monkeypatch.setattr(
        docx_reporter, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    )

    reporter = DocxTrackChangesReporter()
    output = reporter.generate(result, str(tmp_path / "out.docx"))
    assert Path(output).suffix == ".html"
    assert Path(output).exists()
    assert fake_word.quit_called is True


def test_docx_reporter_is_available_true_false(monkeypatch: pytest.MonkeyPatch) -> None:
    fake_word = FakeWord()

    def fake_dispatch(_):
        return fake_word

    monkeypatch.setattr(
        docx_reporter, "win32com", SimpleNamespace(client=SimpleNamespace(Dispatch=fake_dispatch))
    )
    monkeypatch.setattr(
        docx_reporter, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    )

    reporter = DocxTrackChangesReporter()
    assert reporter.is_available() is True

    def raise_dispatch(_):
        raise RuntimeError("no word")

    monkeypatch.setattr(
        docx_reporter, "win32com", SimpleNamespace(client=SimpleNamespace(Dispatch=raise_dispatch))
    )
    reporter.startup_timeout = 0.0
    assert reporter.is_available() is False


def test_docx_reporter_fallback_when_unavailable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    result = make_result(tmp_path)
    reporter = DocxTrackChangesReporter()
    monkeypatch.setattr(docx_reporter, "win32com", None)

    output = reporter.generate(result, str(tmp_path / "report.docx"))
    assert Path(output).suffix == ".html"
    assert Path(output).exists()
    assert not (tmp_path / "report.xlsx").exists()
