from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QStackedWidget,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from core.models import UnsupportedFormatError
from core.registry import ParserRegistry
from ui.comparison_worker import ComparisonWorker
from ui.file_drop_zone import FileDropZone


class MainWindow(QMainWindow):
    MODE_FILE = "file"
    MODE_BATCH = "batch"
    MODE_VERSIONS = "versions"

    def __init__(self) -> None:
        super().__init__()
        ParserRegistry.discover()
        self.supported_extensions = ParserRegistry.supported_extensions()

        self.current_mode = self.MODE_FILE
        self.worker: ComparisonWorker | None = None
        self.last_html_report: str | None = None
        self.last_excel_report: str | None = None

        self.file_a_path = ""
        self.file_b_path = ""
        self.folder_a_path = ""
        self.folder_b_path = ""

        self.setWindowTitle("Change Tracker")
        self.setMinimumSize(800, 500)
        self.setStyleSheet(self._build_styles())

        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(20, 18, 20, 18)
        root_layout.setSpacing(14)

        root_layout.addLayout(self._build_mode_selector())
        root_layout.addWidget(self._build_mode_stack())
        root_layout.addWidget(self._build_bottom_panel())

        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Ready")

        self._set_mode(self.MODE_FILE)
        self._update_action_state()

    def _build_mode_selector(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(8)

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)

        self.file_mode_btn = self._make_mode_button("File vs File", self.MODE_FILE)
        self.batch_mode_btn = self._make_mode_button(
            "Batch (Folder vs Folder)", self.MODE_BATCH
        )
        self.versions_mode_btn = self._make_mode_button(
            "Multi-Version", self.MODE_VERSIONS
        )

        layout.addWidget(self.file_mode_btn)
        layout.addWidget(self.batch_mode_btn)
        layout.addWidget(self.versions_mode_btn)
        layout.addStretch(1)
        return layout

    def _build_mode_stack(self) -> QStackedWidget:
        self.mode_stack = QStackedWidget(self)
        self.file_page = self._build_file_mode_page()
        self.batch_page = self._build_batch_mode_page()
        self.versions_page = self._build_versions_mode_page()
        self.mode_stack.addWidget(self.file_page)
        self.mode_stack.addWidget(self.batch_page)
        self.mode_stack.addWidget(self.versions_page)
        return self.mode_stack

    def _build_file_mode_page(self) -> QWidget:
        page = QWidget(self)
        layout = QHBoxLayout(page)
        layout.setSpacing(12)

        self.file_a_zone, file_a_browse = self._build_file_selector(
            "File A",
            accept_directories=False,
        )
        self.file_b_zone, file_b_browse = self._build_file_selector(
            "File B",
            accept_directories=False,
        )
        self.file_a_zone.file_dropped.connect(self._set_file_a_path)
        self.file_b_zone.file_dropped.connect(self._set_file_b_path)
        file_a_browse.clicked.connect(self._browse_file_a)
        file_b_browse.clicked.connect(self._browse_file_b)

        layout.addWidget(self._wrap_selector("File A", self.file_a_zone, file_a_browse))
        layout.addWidget(self._wrap_selector("File B", self.file_b_zone, file_b_browse))
        return page

    def _build_batch_mode_page(self) -> QWidget:
        page = QWidget(self)
        layout = QHBoxLayout(page)
        layout.setSpacing(12)

        self.folder_a_zone, folder_a_browse = self._build_file_selector(
            "Folder A",
            accept_directories=True,
        )
        self.folder_b_zone, folder_b_browse = self._build_file_selector(
            "Folder B",
            accept_directories=True,
        )
        self.folder_a_zone.file_dropped.connect(self._set_folder_a_path)
        self.folder_b_zone.file_dropped.connect(self._set_folder_b_path)
        folder_a_browse.clicked.connect(self._browse_folder_a)
        folder_b_browse.clicked.connect(self._browse_folder_b)

        layout.addWidget(
            self._wrap_selector("Folder A", self.folder_a_zone, folder_a_browse)
        )
        layout.addWidget(
            self._wrap_selector("Folder B", self.folder_b_zone, folder_b_browse)
        )
        return page

    def _build_versions_mode_page(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(10)

        title = QLabel("Version files (drag to reorder)")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        self.version_list = QListWidget(self)
        self.version_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self.version_list.setDragDropMode(
            QAbstractItemView.DragDropMode.InternalMove
        )
        self.version_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.version_list.setAlternatingRowColors(True)
        self.version_list.model().rowsInserted.connect(self._update_action_state)
        self.version_list.model().rowsRemoved.connect(self._update_action_state)
        self.version_list.model().rowsMoved.connect(self._update_action_state)
        layout.addWidget(self.version_list, 1)

        buttons = QHBoxLayout()
        self.add_versions_btn = QPushButton("Add Files")
        self.remove_versions_btn = QPushButton("Remove Selected")
        self.add_versions_btn.clicked.connect(self._add_version_files)
        self.remove_versions_btn.clicked.connect(self._remove_selected_versions)
        buttons.addWidget(self.add_versions_btn)
        buttons.addWidget(self.remove_versions_btn)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        return page

    def _build_bottom_panel(self) -> QWidget:
        panel = QFrame(self)
        panel.setObjectName("bottomPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        output_row = QHBoxLayout()
        output_label = QLabel("Output folder:")
        self.output_line = QLineEdit("./output/")
        self.output_line.textChanged.connect(self._update_action_state)
        browse_output_btn = QPushButton("Browse")
        browse_output_btn.clicked.connect(self._browse_output_folder)
        output_row.addWidget(output_label)
        output_row.addWidget(self.output_line, 1)
        output_row.addWidget(browse_output_btn)
        layout.addLayout(output_row)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        links = QHBoxLayout()
        self.open_html_btn = QPushButton("Open HTML Report")
        self.open_excel_btn = QPushButton("Open Excel Report")
        self.open_html_btn.clicked.connect(lambda: self._open_report(self.last_html_report))
        self.open_excel_btn.clicked.connect(
            lambda: self._open_report(self.last_excel_report)
        )
        self.open_html_btn.setVisible(False)
        self.open_excel_btn.setVisible(False)
        links.addWidget(self.open_html_btn)
        links.addWidget(self.open_excel_btn)
        links.addStretch(1)
        layout.addLayout(links)

        self.compare_btn = QPushButton("Compare")
        self.compare_btn.setObjectName("compareButton")
        self.compare_btn.clicked.connect(self._start_comparison)
        layout.addWidget(self.compare_btn)
        return panel

    def _build_file_selector(
        self, title: str, *, accept_directories: bool
    ) -> tuple[FileDropZone, QPushButton]:
        zone = FileDropZone(
            title,
            accept_directories=accept_directories,
            allowed_extensions=None if accept_directories else self.supported_extensions,
        )
        browse = QPushButton("Browse")
        return zone, browse

    def _wrap_selector(
        self, title: str, zone: FileDropZone, browse_btn: QPushButton
    ) -> QWidget:
        wrapper = QFrame(self)
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        layout.addWidget(label)
        layout.addWidget(zone)
        layout.addWidget(browse_btn, 0, Qt.AlignmentFlag.AlignLeft)
        return wrapper

    def _make_mode_button(self, text: str, mode: str) -> QPushButton:
        button = QPushButton(text)
        button.setCheckable(True)
        button.clicked.connect(lambda: self._set_mode(mode))
        self.mode_group.addButton(button)
        return button

    def _set_mode(self, mode: str) -> None:
        self.current_mode = mode
        if mode == self.MODE_FILE:
            self.mode_stack.setCurrentWidget(self.file_page)
            self.compare_btn.setText("Compare")
            self.file_mode_btn.setChecked(True)
        elif mode == self.MODE_BATCH:
            self.mode_stack.setCurrentWidget(self.batch_page)
            self.compare_btn.setText("Compare All")
            self.batch_mode_btn.setChecked(True)
        else:
            self.mode_stack.setCurrentWidget(self.versions_page)
            self.compare_btn.setText("Compare Versions")
            self.versions_mode_btn.setChecked(True)
        self._update_action_state()

    def _browse_file_a(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File A",
            "",
            self._supported_filter(),
        )
        if path:
            self._set_file_a_path(path)

    def _browse_file_b(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File B",
            "",
            self._supported_filter(),
        )
        if path:
            self._set_file_b_path(path)

    def _browse_folder_a(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Folder A")
        if path:
            self._set_folder_a_path(path)

    def _browse_folder_b(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Folder B")
        if path:
            self._set_folder_b_path(path)

    def _browse_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if path:
            self.output_line.setText(path)

    def _set_file_a_path(self, path: str) -> None:
        self.file_a_path = path
        self.file_a_zone.set_path(path)
        self._update_action_state()

    def _set_file_b_path(self, path: str) -> None:
        self.file_b_path = path
        self.file_b_zone.set_path(path)
        self._update_action_state()

    def _set_folder_a_path(self, path: str) -> None:
        self.folder_a_path = path
        self.folder_a_zone.set_path(path)
        self._update_action_state()

    def _set_folder_b_path(self, path: str) -> None:
        self.folder_b_path = path
        self.folder_b_zone.set_path(path)
        self._update_action_state()

    def _add_version_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Version Files",
            "",
            self._supported_filter(),
        )
        for path in paths:
            self.version_list.addItem(QListWidgetItem(path))
        self._update_action_state()

    def _remove_selected_versions(self) -> None:
        selected = self.version_list.selectedItems()
        for item in selected:
            row = self.version_list.row(item)
            self.version_list.takeItem(row)
        self._update_action_state()

    def _start_comparison(self) -> None:
        output_dir = self.output_line.text().strip()
        if not output_dir:
            return

        payload: dict[str, object]
        if self.current_mode == self.MODE_FILE:
            if not self.file_a_path or not self.file_b_path:
                return
            ext_a = Path(self.file_a_path).suffix.lower()
            ext_b = Path(self.file_b_path).suffix.lower()
            if ext_a != ext_b:
                QMessageBox.warning(
                    self,
                    "Invalid input",
                    "Files must have the same extension.",
                )
                return
            payload = {
                "file_a": self.file_a_path,
                "file_b": self.file_b_path,
                "output_dir": output_dir,
            }
        elif self.current_mode == self.MODE_BATCH:
            if not self.folder_a_path or not self.folder_b_path:
                return
            payload = {
                "folder_a": self.folder_a_path,
                "folder_b": self.folder_b_path,
                "output_dir": output_dir,
            }
        else:
            files = [self.version_list.item(i).text() for i in range(self.version_list.count())]
            if len(files) < 2:
                return
            payload = {"files": files, "output_dir": output_dir}

        self.last_html_report = None
        self.last_excel_report = None
        self.open_html_btn.setVisible(False)
        self.open_excel_btn.setVisible(False)
        self.compare_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Starting comparison...")

        self.worker = ComparisonWorker(self.current_mode, payload, self)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)
        self.worker.start()

    def _on_worker_progress(self, message: str, percent: float) -> None:
        value = max(0, min(100, int(percent * 100)))
        self.progress_bar.setValue(value)
        self.statusBar().showMessage(message)

    def _on_worker_finished(self, payload: dict) -> None:
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.worker = None

        mode = payload.get("mode")
        if mode == self.MODE_FILE:
            outputs = [str(path) for path in payload.get("outputs", [])]
            self.last_html_report = next(
                (path for path in outputs if path.lower().endswith(".html")), None
            )
            self.last_excel_report = next(
                (path for path in outputs if path.lower().endswith(".xlsx")), None
            )
            comparison = payload.get("comparison")
            if comparison is not None:
                stats = comparison.statistics
                self.statusBar().showMessage(
                    "Done: total={total}, added={added}, deleted={deleted}, "
                    "modified={modified}, unchanged={unchanged}".format(
                        total=stats.total_segments,
                        added=stats.added,
                        deleted=stats.deleted,
                        modified=stats.modified,
                        unchanged=stats.unchanged,
                    )
                )
            if (
                Path(self.file_a_path).suffix.lower() == ".docx"
                and not any(path.lower().endswith(".docx") for path in outputs)
            ):
                QMessageBox.warning(
                    self,
                    "Word unavailable",
                    "Microsoft Word not found, Track Changes unavailable. "
                    "HTML and Excel reports will be generated.",
                )
        elif mode == self.MODE_BATCH:
            result = payload["result"]
            self.last_html_report = result.summary_report_path
            self.last_excel_report = result.summary_excel_path
            self.statusBar().showMessage(
                "Done: total={total}, compared={compared}, only_in_a={only_a}, "
                "only_in_b={only_b}, errors={errors}".format(
                    total=result.total_files,
                    compared=result.compared_files,
                    only_a=result.only_in_a,
                    only_b=result.only_in_b,
                    errors=result.errors,
                )
            )
        elif mode == self.MODE_VERSIONS:
            result = payload["result"]
            self.last_html_report = result.summary_report_path
            self.last_excel_report = None
            self.statusBar().showMessage(
                f"Done: {len(result.comparisons)} comparisons generated"
            )

        self.open_html_btn.setVisible(bool(self.last_html_report))
        self.open_excel_btn.setVisible(bool(self.last_excel_report))
        self._update_action_state()

    def _on_worker_error(self, message: str) -> None:
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.worker = None
        self.statusBar().showMessage("Failed")

        if "Unsupported format" in message or isinstance(
            message, UnsupportedFormatError
        ):
            QMessageBox.warning(self, "Unsupported format", str(message))
            return
        QMessageBox.critical(self, "Comparison error", str(message))

    def _open_report(self, path: str | None) -> None:
        if not path:
            return
        if not Path(path).exists():
            QMessageBox.warning(self, "File not found", f"Report not found:\n{path}")
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "Open failed", str(exc))

    def _update_action_state(self) -> None:
        output_ok = bool(self.output_line.text().strip())
        enabled = False
        if self.current_mode == self.MODE_FILE:
            enabled = bool(self.file_a_path and self.file_b_path and output_ok)
        elif self.current_mode == self.MODE_BATCH:
            enabled = bool(self.folder_a_path and self.folder_b_path and output_ok)
        elif self.current_mode == self.MODE_VERSIONS:
            enabled = bool(self.version_list.count() >= 2 and output_ok)
        self.compare_btn.setEnabled(enabled and self.worker is None)

    def _supported_filter(self) -> str:
        if not self.supported_extensions:
            return "All files (*.*)"
        patterns = " ".join(f"*{ext}" for ext in self.supported_extensions)
        return f"Supported files ({patterns});;All files (*.*)"

    @staticmethod
    def _build_styles() -> str:
        return """
QWidget {
  background: #f6f7fb;
  color: #1f2933;
  font-family: "Segoe UI";
  font-size: 13px;
}
QPushButton {
  background: #ffffff;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  padding: 7px 12px;
}
QPushButton:checked {
  background: #111827;
  color: #ffffff;
  border-color: #111827;
}
QPushButton#compareButton {
  background: #1d4ed8;
  color: #ffffff;
  border-color: #1d4ed8;
  font-weight: 600;
}
QPushButton:disabled {
  background: #e2e8f0;
  color: #94a3b8;
  border-color: #cbd5e1;
}
QLineEdit, QListWidget {
  background: #ffffff;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  padding: 6px 8px;
}
QFrame#bottomPanel {
  background: #ffffff;
  border: 1px solid #e2e8f0;
  border-radius: 12px;
}
QLabel#sectionTitle {
  font-weight: 600;
  color: #334155;
}
QProgressBar {
  background: #e2e8f0;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  text-align: center;
}
QProgressBar::chunk {
  background: #1d4ed8;
  border-radius: 7px;
}
"""


def run_gui() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

