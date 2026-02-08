from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from ui.comparison_worker import ComparisonWorker
from ui.file_drop_zone import FileDropZone
from ui.main_window import MainWindow


@pytest.fixture(scope="session")
def app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_main_window_creation(app: QApplication) -> None:
    window = MainWindow()
    assert window.windowTitle() == "Change Tracker"
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
    window.close()
