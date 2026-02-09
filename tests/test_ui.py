from __future__ import annotations

from datetime import datetime, timezone
import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from core.models import ChangeStatistics, ComparisonResult, ParsedDocument
from ui.comparison_worker import ComparisonWorker
from ui.file_drop_zone import FileDropZone
from ui.main_window import MainWindow

FIXTURES = Path(__file__).resolve().parent / "fixtures"


@pytest.fixture(scope="session")
def app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_main_window_creation(app: QApplication) -> None:
    window = MainWindow()
    assert window.windowTitle() == "Diff View"
    assert window.minimumWidth() == 800
    assert window.minimumHeight() == 500
    window.close()


def test_file_drop_zone_creation(app: QApplication) -> None:
    zone = FileDropZone("Test Zone", accept_directories=False, allowed_extensions=[".txt"])
    assert zone.accept_directories is False
    assert ".txt" in zone.allowed_extensions
    zone.close()


def test_comparison_worker_creation() -> None:
    worker = ComparisonWorker("file", {"pairs": [("a", "b")], "output_dir": "out"})
    assert worker.mode == "file"
    assert "pairs" in worker.payload


def test_comparison_worker_pair_folder_name_is_compact() -> None:
    file_a = ("a" * 150) + ".txt"
    file_b = ("b" * 150) + ".txt"
    folder = ComparisonWorker._pair_folder_name(1, file_a, file_b)
    assert folder.startswith("001_")
    assert "_vs_" in folder
    assert len(folder) <= 100


def test_main_window_mode_switching(app: QApplication) -> None:
    window = MainWindow()

    window._set_mode(window.MODE_FILE)
    assert window.current_mode == window.MODE_FILE
    assert window.compare_btn.text() == "Compare"

    window._set_mode(window.MODE_VERSIONS)
    assert window.current_mode == window.MODE_VERSIONS
    assert window.compare_btn.text() == "Compare Versions"

    window._set_mode(window.MODE_QA_VERIFY)
    assert window.current_mode == window.MODE_QA_VERIFY
    assert window.compare_btn.text() == "Verify QA"
    window.close()


def test_main_window_excel_column_validation() -> None:
    assert MainWindow._normalize_excel_column_input(" a ") == "A"
    assert MainWindow._normalize_excel_column_input("12") == "12"
    assert MainWindow._normalize_excel_column_input("") is None
    with pytest.raises(ValueError):
        MainWindow._normalize_excel_column_input("A-1")


def test_main_window_has_help_and_about_menu(app: QApplication) -> None:
    window = MainWindow()
    top_actions = [action.text() for action in window.menuBar().actions()]
    assert "Справка" in top_actions
    assert "О программе" in top_actions
    window.close()


def test_excel_source_row_visibility_for_excel_files(app: QApplication) -> None:
    window = MainWindow()
    assert window.excel_source_options_widget.isHidden() is True

    window.file_a_zone.add_files([str(FIXTURES / "sample_a.xlsx")])
    assert window.excel_source_options_widget.isHidden() is False

    window.file_a_zone.clear_files()
    assert window.excel_source_options_widget.isHidden() is True
    window.close()


def test_main_window_shows_no_changes_message_for_empty_report(
    app: QApplication, tmp_path: Path, monkeypatch
) -> None:
    window = MainWindow()
    report_html = tmp_path / "report.html"
    report_excel = tmp_path / "report.xlsx"
    report_html.write_text("<html></html>", encoding="utf-8")
    report_excel.write_bytes(b"")

    doc = ParsedDocument(segments=[], format_name="TXT", file_path="a.txt")
    comparison = ComparisonResult(
        file_a=doc,
        file_b=doc,
        changes=[],
        statistics=ChangeStatistics.from_changes([]),
        timestamp=datetime.now(timezone.utc),
    )

    messages: list[str] = []

    def fake_information(_parent, _title, text):
        messages.append(text)
        return 0

    monkeypatch.setattr("ui.main_window.QMessageBox.information", fake_information)

    window._on_worker_finished(
        {
            "mode": window.MODE_FILE,
            "multi": False,
            "outputs": [str(report_html), str(report_excel)],
            "comparison": comparison,
        }
    )

    assert any("Правок не найдено" in message for message in messages)
    window.close()
