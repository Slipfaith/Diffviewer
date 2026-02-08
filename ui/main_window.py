from __future__ import annotations

from collections import defaultdict
import os
from pathlib import Path
import subprocess
import sys
import webbrowser

from PyQt6.QtCore import QEvent, Qt, QUrl, pyqtSignal
from PyQt6.QtGui import QDesktopServices, QDragEnterEvent, QDropEvent, QIcon
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
from ui.file_tile_drop_zone import FileTileDropZone, TileVisualState


def _resolve_app_icon() -> QIcon | None:
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(str(meipass)) / "Diffviewer.ico")
        candidates.append(Path(sys.executable).resolve().parent / "Diffviewer.ico")
    candidates.append(Path(__file__).resolve().parents[1] / "Diffviewer.ico")

    for icon_path in candidates:
        try:
            if not icon_path.exists():
                continue
            icon = QIcon(str(icon_path))
            if not icon.isNull():
                return icon
        except Exception:
            continue
    return None


def _set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        app_id = "Diffviewer.ChangeTracker"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


class VersionFileListWidget(QListWidget):
    files_dropped = pyqtSignal(list)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.source() is self:
            super().dragEnterEvent(event)
            return
        if self.extract_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if event.source() is self:
            super().dragMoveEvent(event)
            return
        if self.extract_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        if event.source() is self:
            super().dropEvent(event)
            return
        paths = self.extract_paths_from_mime(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)

    @staticmethod
    def extract_paths_from_mime(mime_data) -> list[str]:
        candidates: list[str] = []

        if mime_data.hasUrls():
            for url in mime_data.urls():
                local = url.toLocalFile()
                if local:
                    candidates.append(local)
        elif mime_data.hasText():
            candidates.extend(
                line.strip()
                for line in mime_data.text().splitlines()
                if line.strip()
            )

        paths: list[str] = []
        for raw in candidates:
            path = Path(raw)
            if not path.exists() or not path.is_file():
                continue
            try:
                resolved = str(path.resolve())
            except Exception:
                resolved = str(path)
            paths.append(resolved)
        return paths


class MainWindow(QMainWindow):
    MODE_FILE = "file"
    MODE_VERSIONS = "versions"

    def __init__(self) -> None:
        super().__init__()
        ParserRegistry.discover()
        self.supported_extensions = ParserRegistry.supported_extensions()

        self.current_mode = self.MODE_FILE
        self.worker: ComparisonWorker | None = None
        self.last_html_report: str | None = None
        self.last_excel_report: str | None = None

        self.manual_file_pairs: dict[str, str] = {}
        self.pending_file_a: str | None = None

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
        self.versions_mode_btn = self._make_mode_button(
            "Multi-Version", self.MODE_VERSIONS
        )
        self.versions_mode_btn.setAcceptDrops(True)
        self.versions_mode_btn.installEventFilter(self)

        layout.addWidget(self.file_mode_btn)
        layout.addWidget(self.versions_mode_btn)
        layout.addStretch(1)
        return layout

    def _build_mode_stack(self) -> QStackedWidget:
        self.mode_stack = QStackedWidget(self)
        self.file_page = self._build_file_mode_page()
        self.versions_page = self._build_versions_mode_page()
        self.mode_stack.addWidget(self.file_page)
        self.mode_stack.addWidget(self.versions_page)
        return self.mode_stack

    def _build_file_mode_page(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(10)

        lists_row = QHBoxLayout()
        lists_row.setSpacing(12)

        self.file_a_zone = FileTileDropZone(
            "File A",
            allowed_extensions=self.supported_extensions,
            parent=self,
        )
        self.file_b_zone = FileTileDropZone(
            "File B",
            allowed_extensions=self.supported_extensions,
            parent=self,
        )
        self.file_a_zone.files_changed.connect(self._on_file_lists_changed)
        self.file_b_zone.files_changed.connect(self._on_file_lists_changed)
        self.file_a_zone.file_left_clicked.connect(self._on_file_a_tile_clicked)
        self.file_b_zone.file_left_clicked.connect(self._on_file_b_tile_clicked)

        lists_row.addWidget(self.file_a_zone)
        lists_row.addWidget(self.file_b_zone)
        layout.addLayout(lists_row, 1)

        controls = QHBoxLayout()
        self.clear_file_lists_btn = QPushButton("Clear lists")
        self.clear_file_lists_btn.clicked.connect(self._clear_file_lists)
        controls.addStretch(1)
        controls.addWidget(self.clear_file_lists_btn)
        layout.addLayout(controls)

        self._refresh_file_pairing_visuals()
        return page

    def _build_versions_mode_page(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(10)

        title = QLabel("Version files (drag to reorder)")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        self.version_list = VersionFileListWidget(self)
        self.version_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self.version_list.setDragDropMode(
            QAbstractItemView.DragDropMode.DragDrop
        )
        self.version_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.version_list.setAlternatingRowColors(True)
        self.version_list.files_dropped.connect(self._add_version_paths)
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
        self.compare_btn.setMinimumWidth(170)
        self.compare_btn.setMaximumWidth(240)
        action_row = QHBoxLayout()
        action_row.addStretch(1)
        action_row.addWidget(self.compare_btn)
        action_row.addStretch(1)
        layout.addLayout(action_row)
        return panel

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
            self._refresh_file_pairing_visuals()
        else:
            self.mode_stack.setCurrentWidget(self.versions_page)
            self.compare_btn.setText("Compare Versions")
            self.versions_mode_btn.setChecked(True)
        self._update_action_state()

    def eventFilter(self, watched, event) -> bool:  # type: ignore[override]
        if watched is self.versions_mode_btn:
            if event.type() in {QEvent.Type.DragEnter, QEvent.Type.DragMove}:
                paths = VersionFileListWidget.extract_paths_from_mime(event.mimeData())
                if paths:
                    event.acceptProposedAction()
                    return True
            elif event.type() == QEvent.Type.Drop:
                paths = VersionFileListWidget.extract_paths_from_mime(event.mimeData())
                if paths:
                    self._set_mode(self.MODE_VERSIONS)
                    self._add_version_paths(paths)
                    event.acceptProposedAction()
                    return True
        return super().eventFilter(watched, event)

    def _on_file_lists_changed(self, _paths: list[str]) -> None:
        self._cleanup_file_pair_state()
        self._refresh_file_pairing_visuals()
        self._update_action_state()

    def _on_file_a_tile_clicked(self, file_path: str) -> None:
        self.pending_file_a = file_path
        self._refresh_file_pairing_visuals()
        self._update_action_state()

    def _on_file_b_tile_clicked(self, file_path: str) -> None:
        if self.pending_file_a is None:
            return
        self._set_manual_file_pair(self.pending_file_a, file_path)
        self.pending_file_a = None
        self._refresh_file_pairing_visuals()
        self._update_action_state()

    def _clear_file_lists(self) -> None:
        self.file_a_zone.clear_files()
        self.file_b_zone.clear_files()
        self.manual_file_pairs.clear()
        self.pending_file_a = None
        self._refresh_file_pairing_visuals()
        self._update_action_state()

    def _cleanup_file_pair_state(self) -> None:
        files_a = set(self.file_a_zone.file_paths())
        files_b = set(self.file_b_zone.file_paths())
        self.manual_file_pairs = {
            file_a: file_b
            for file_a, file_b in self.manual_file_pairs.items()
            if file_a in files_a and file_b in files_b
        }
        if self.pending_file_a not in files_a:
            self.pending_file_a = None

    def _auto_file_pairs(
        self, available_a: list[str], available_b: list[str]
    ) -> dict[str, str]:
        grouped_b: dict[str, list[str]] = defaultdict(list)
        for file_b in available_b:
            grouped_b[Path(file_b).name.casefold()].append(file_b)

        pairs: dict[str, str] = {}
        for file_a in available_a:
            key = Path(file_a).name.casefold()
            bucket = grouped_b.get(key)
            if bucket:
                pairs[file_a] = bucket.pop(0)
        return pairs

    def _current_file_pairs_map(self) -> dict[str, str]:
        files_a = self.file_a_zone.file_paths()
        files_b = self.file_b_zone.file_paths()
        if not files_a or not files_b:
            return {}

        self._cleanup_file_pair_state()

        used_a = set(self.manual_file_pairs.keys())
        used_b = set(self.manual_file_pairs.values())
        remaining_a = [path for path in files_a if path not in used_a]
        remaining_b = [path for path in files_b if path not in used_b]

        auto_pairs = self._auto_file_pairs(remaining_a, remaining_b)
        combined = dict(self.manual_file_pairs)
        combined.update(auto_pairs)
        return combined

    def _ordered_file_pairs(self) -> list[tuple[str, str]]:
        files_a = self.file_a_zone.file_paths()
        pairs_map = self._current_file_pairs_map()
        return [(file_a, pairs_map[file_a]) for file_a in files_a if file_a in pairs_map]

    def _set_manual_file_pair(self, file_a: str, file_b: str) -> None:
        cleaned: dict[str, str] = {}
        for key_a, key_b in self.manual_file_pairs.items():
            if key_a == file_a or key_b == file_b:
                continue
            cleaned[key_a] = key_b
        cleaned[file_a] = file_b
        self.manual_file_pairs = cleaned

    def _refresh_file_pairing_visuals(self) -> None:
        files_a = self.file_a_zone.file_paths()
        files_b = self.file_b_zone.file_paths()
        active = bool(files_a and files_b)
        pairs_map = self._current_file_pairs_map() if active else {}
        matched_b = set(pairs_map.values())

        states_a: dict[str, TileVisualState] = {}
        states_b: dict[str, TileVisualState] = {}

        for file_a in files_a:
            is_matched = file_a in pairs_map
            states_a[file_a] = TileVisualState(
                matched=is_matched,
                unmatched=active and not is_matched,
                selected=file_a == self.pending_file_a,
            )

        show_candidates = active and self.pending_file_a is not None
        for file_b in files_b:
            is_matched = file_b in matched_b
            states_b[file_b] = TileVisualState(
                matched=is_matched,
                unmatched=active and not is_matched,
                candidate=show_candidates,
            )

        self.file_a_zone.apply_states(states_a)
        self.file_b_zone.apply_states(states_b)

    def _browse_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if path:
            self.output_line.setText(path)

    def _add_version_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Version Files",
            "",
            self._supported_filter(),
        )
        self._add_version_paths(paths)

    def _add_version_paths(self, paths: list[str]) -> None:
        existing = {self.version_list.item(i).text() for i in range(self.version_list.count())}
        added = False
        for path in paths:
            candidate = Path(path)
            if not candidate.exists() or not candidate.is_file():
                continue
            if self.supported_extensions and candidate.suffix.lower() not in self.supported_extensions:
                continue
            try:
                normalized = str(candidate.resolve())
            except Exception:
                normalized = str(candidate)
            if normalized in existing:
                continue
            self.version_list.addItem(QListWidgetItem(normalized))
            existing.add(normalized)
            added = True
        if added:
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
            pairs = self._ordered_file_pairs()
            if not pairs:
                return
            mismatched = [
                (file_a, file_b)
                for file_a, file_b in pairs
                if Path(file_a).suffix.lower() != Path(file_b).suffix.lower()
            ]
            if mismatched:
                bad_a, bad_b = mismatched[0]
                QMessageBox.warning(
                    self,
                    "Invalid input",
                    "Mapped files must have the same extension:\n"
                    f"{Path(bad_a).name} vs {Path(bad_b).name}",
                )
                return
            payload = {
                "pairs": pairs,
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
            if payload.get("multi"):
                file_results = list(payload.get("file_results", []))
                successful = [item for item in file_results if not item.get("error")]
                failed = [item for item in file_results if item.get("error")]

                self.last_html_report = None
                self.last_excel_report = None
                if successful:
                    first_outputs = [str(path) for path in successful[0].get("outputs", [])]
                    self.last_html_report = next(
                        (path for path in first_outputs if path.lower().endswith(".html")),
                        None,
                    )
                    self.last_excel_report = next(
                        (path for path in first_outputs if path.lower().endswith(".xlsx")),
                        None,
                    )

                self.statusBar().showMessage(
                    "Done: compared={ok}, errors={err}, pairs={total}".format(
                        ok=len(successful),
                        err=len(failed),
                        total=len(file_results),
                    )
                )
                if failed:
                    preview = "\n".join(
                        f"- {Path(item['file_a']).name} vs {Path(item['file_b']).name}: {item['error']}"
                        for item in failed[:5]
                    )
                    QMessageBox.warning(
                        self,
                        "Some comparisons failed",
                        preview,
                    )
            else:
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
        report_path = Path(path).expanduser()
        try:
            report_path = report_path.resolve(strict=False)
        except Exception:
            pass

        if not report_path.exists():
            QMessageBox.warning(
                self,
                "File not found",
                f"Report not found:\n{report_path}",
            )
            return

        if report_path.suffix.lower() in {".html", ".htm"}:
            try:
                if webbrowser.open_new_tab(report_path.as_uri()):
                    return
            except Exception:
                pass

        if QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path))):
            return

        try:
            os.startfile(str(report_path))  # type: ignore[attr-defined]
        except Exception as exc:
            try:
                subprocess.Popen(["explorer", str(report_path)])
                return
            except Exception:
                QMessageBox.warning(
                    self,
                    "Open failed",
                    f"Cannot open report:\n{report_path}\n\n{exc}",
                )

    def _update_action_state(self) -> None:
        output_ok = bool(self.output_line.text().strip())
        enabled = False
        if self.current_mode == self.MODE_FILE:
            enabled = bool(self._ordered_file_pairs() and output_ok)
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
QFrame#fileTileDropZone {
  background: #f8fafc;
  border: 1px solid #dbe4f0;
  border-radius: 12px;
}
QLabel#fileTileDropZoneTitle {
  color: #334155;
  font-weight: 600;
  font-size: 12px;
}
QListWidget#fileTileList {
  background: #ffffff;
  border: 1px solid #e2e8f0;
  border-radius: 10px;
  padding: 6px;
}
QLabel#fileTileHint {
  color: #64748b;
  font-size: 12px;
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
    _set_windows_app_id()
    app = QApplication.instance() or QApplication([])
    icon = _resolve_app_icon()
    if icon is not None:
        app.setWindowIcon(icon)
    window = MainWindow()
    if icon is not None:
        window.setWindowIcon(icon)
    window.show()
    app.exec()
