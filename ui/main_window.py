from __future__ import annotations

from collections import defaultdict
from datetime import datetime
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
from core.qa_verify import (
    QAScanResult,
    QASheetConfig,
    QAVerificationResult,
    QAVerifier,
    STATUS_APPLIED,
    STATUS_CANNOT_VERIFY,
    STATUS_NOT_APPLICABLE,
    STATUS_NOT_APPLIED,
)
from core.registry import ParserRegistry
from ui.comparison_worker import ComparisonWorker
from ui.file_tile_drop_zone import FileTileDropZone, TileVisualState
from ui.qa_column_mapping_dialog import QAColumnMappingDialog


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


def _resolve_app_version() -> str:
    main_module = sys.modules.get("__main__")
    raw_version = getattr(main_module, "APP_VERSION", None)
    if isinstance(raw_version, str) and raw_version.strip():
        return raw_version.strip()
    return "1.0"


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
    MODE_QA_VERIFY = "qa_verify"

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
        self.qa_sheet_configs: list[QASheetConfig] = []
        self.qa_scan_warnings: list[str] = []
        self.qa_result: QAVerificationResult | None = None

        self.setWindowTitle("Diff View")
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

        self._build_top_menu()
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Ready")

        self._set_mode(self.MODE_FILE)
        self._update_action_state()

    def _build_top_menu(self) -> None:
        menu_bar = self.menuBar()
        help_action = menu_bar.addAction("Справка")
        about_action = menu_bar.addAction("О программе")
        help_action.triggered.connect(self._show_help_dialog)
        about_action.triggered.connect(self._show_about_dialog)

    def _show_help_dialog(self) -> None:
        text = (
            "Diff View помогает сравнивать файлы и проверять изменения переводов.\n\n"
            "Основные возможности:\n"
            "1. File vs File: сравнение пар файлов с отчетами HTML/XLSX.\n"
            "2. Multi-Version: сводное сравнение нескольких версий файла.\n"
            "3. QA Verify: проверка применения QA-правок TP/FP.\n\n"
            "from Sha by slipfaith."
        )
        QMessageBox.information(self, "Справка", text)

    def _show_about_dialog(self) -> None:
        QMessageBox.information(
            self,
            "О программе",
            f"Diff View\nВерсия: {_resolve_app_version()}",
        )

    def _build_mode_selector(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(8)

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)

        self.file_mode_btn = self._make_mode_button("File vs File", self.MODE_FILE)
        self.versions_mode_btn = self._make_mode_button(
            "Multi-Version", self.MODE_VERSIONS
        )
        self.qa_verify_mode_btn = self._make_mode_button(
            "QA Verify", self.MODE_QA_VERIFY
        )
        self.versions_mode_btn.setAcceptDrops(True)
        self.versions_mode_btn.installEventFilter(self)

        layout.addWidget(self.file_mode_btn)
        layout.addWidget(self.versions_mode_btn)
        layout.addWidget(self.qa_verify_mode_btn)
        layout.addStretch(1)
        return layout

    def _build_mode_stack(self) -> QStackedWidget:
        self.mode_stack = QStackedWidget(self)
        self.file_page = self._build_file_mode_page()
        self.versions_page = self._build_versions_mode_page()
        self.qa_verify_page = self._build_qa_verify_page()
        self.mode_stack.addWidget(self.file_page)
        self.mode_stack.addWidget(self.versions_page)
        self.mode_stack.addWidget(self.qa_verify_page)
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

        self.excel_source_options_widget = QWidget(self)
        excel_options = QHBoxLayout(self.excel_source_options_widget)
        excel_options.setContentsMargins(0, 0, 0, 0)
        excel_options.setSpacing(8)
        excel_label = QLabel("Bilingual Excel Source columns (optional):")
        self.excel_source_col_a_input = QLineEdit(self)
        self.excel_source_col_a_input.setPlaceholderText("File A: A")
        self.excel_source_col_a_input.setMaximumWidth(130)
        self.excel_source_col_a_input.textChanged.connect(self._update_action_state)
        self.excel_source_col_b_input = QLineEdit(self)
        self.excel_source_col_b_input.setPlaceholderText("File B: A")
        self.excel_source_col_b_input.setMaximumWidth(130)
        self.excel_source_col_b_input.textChanged.connect(self._update_action_state)
        excel_options.addWidget(excel_label)
        excel_options.addWidget(self.excel_source_col_a_input)
        excel_options.addWidget(self.excel_source_col_b_input)
        excel_options.addStretch(1)
        layout.addWidget(self.excel_source_options_widget)
        self.excel_source_options_widget.setVisible(False)

        self._refresh_file_pairing_visuals()
        self._update_excel_source_controls_visibility()
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

    def _build_qa_verify_page(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(10)

        lists_row = QHBoxLayout()
        lists_row.setSpacing(12)

        self.qa_reports_zone = FileTileDropZone(
            "QA Reports (.xlsx)",
            allowed_extensions=[".xlsx"],
            parent=self,
        )
        self.qa_reports_zone.files_changed.connect(self._on_qa_reports_changed)

        self.qa_final_zone = FileTileDropZone(
            "Final XLIFF Files",
            allowed_extensions=[".xliff", ".xlf", ".sdlxliff", ".mqxliff"],
            parent=self,
        )
        self.qa_final_zone.files_changed.connect(self._on_qa_final_files_changed)

        lists_row.addWidget(self.qa_reports_zone, 1)
        lists_row.addWidget(self.qa_final_zone, 1)
        layout.addLayout(lists_row, 1)

        controls_row = QHBoxLayout()
        controls_row.setSpacing(8)

        self.qa_mapping_status_label = QLabel("Load QA reports to detect columns.")
        self.qa_mapping_status_label.setObjectName("sectionTitle")
        controls_row.addWidget(self.qa_mapping_status_label, 1)

        self.qa_map_columns_btn = QPushButton("Column Mapping")
        self.qa_map_columns_btn.clicked.connect(self._open_qa_mapping_dialog)
        controls_row.addWidget(self.qa_map_columns_btn)

        self.qa_export_btn = QPushButton("Export Results")
        self.qa_export_btn.clicked.connect(self._export_qa_results)
        controls_row.addWidget(self.qa_export_btn)
        layout.addLayout(controls_row)

        summary_row = QHBoxLayout()
        summary_row.setSpacing(8)
        self.qa_summary_labels: dict[str, QLabel] = {}
        for status in (
            STATUS_APPLIED,
            STATUS_NOT_APPLIED,
            STATUS_CANNOT_VERIFY,
            STATUS_NOT_APPLICABLE,
        ):
            box = QFrame(self)
            box.setObjectName("qaSummaryBox")
            box_layout = QHBoxLayout(box)
            box_layout.setContentsMargins(8, 4, 8, 4)
            box_layout.setSpacing(6)
            label = QLabel(status)
            value = QLabel("0")
            value.setObjectName("qaSummaryValue")
            self.qa_summary_labels[status] = value
            box_layout.addWidget(label)
            box_layout.addWidget(value)
            summary_row.addWidget(box)
        summary_row.addStretch(1)
        layout.addLayout(summary_row)

        hint = QLabel(
            "Verification results are shown in summary and exported to Excel.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self._update_qa_controls()
        return page

    def _build_bottom_panel(self) -> QWidget:
        panel = QFrame(self)
        panel.setObjectName("bottomPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        self.output_controls = QWidget(self)
        output_row = QHBoxLayout(self.output_controls)
        output_row.setContentsMargins(0, 0, 0, 0)
        output_row.setSpacing(8)
        self.output_label = QLabel("Output folder:")
        self.output_line = QLineEdit("./output/")
        self.output_line.textChanged.connect(self._update_action_state)
        self.browse_output_btn = QPushButton("Browse")
        self.browse_output_btn.clicked.connect(self._browse_output_folder)
        output_row.addWidget(self.output_label)
        output_row.addWidget(self.output_line, 1)
        output_row.addWidget(self.browse_output_btn)
        layout.addWidget(self.output_controls)

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
        qa_mode = mode == self.MODE_QA_VERIFY
        self.output_controls.setVisible(not qa_mode)
        if mode == self.MODE_FILE:
            self.mode_stack.setCurrentWidget(self.file_page)
            self.compare_btn.setText("Compare")
            self.file_mode_btn.setChecked(True)
            self._refresh_file_pairing_visuals()
        elif mode == self.MODE_VERSIONS:
            self.mode_stack.setCurrentWidget(self.versions_page)
            self.compare_btn.setText("Compare Versions")
            self.versions_mode_btn.setChecked(True)
        else:
            self.mode_stack.setCurrentWidget(self.qa_verify_page)
            self.compare_btn.setText("Verify QA")
            self.qa_verify_mode_btn.setChecked(True)
            self.open_html_btn.setVisible(False)
            self.open_excel_btn.setVisible(False)
            self.last_html_report = None
            self.last_excel_report = None
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
        self._update_excel_source_controls_visibility()
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
        self._update_excel_source_controls_visibility()
        self._update_action_state()

    def _update_excel_source_controls_visibility(self) -> None:
        should_show = any(
            Path(path).suffix.lower() in {".xlsx", ".xls"}
            for path in self.file_a_zone.file_paths() + self.file_b_zone.file_paths()
        )
        self.excel_source_options_widget.setVisible(should_show)

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

    def _on_qa_reports_changed(self, _paths: list[str]) -> None:
        self.qa_result = None
        self._scan_qa_reports()
        self._populate_qa_results_table(None)
        self._update_action_state()

    def _on_qa_final_files_changed(self, _paths: list[str]) -> None:
        self.qa_result = None
        self._update_qa_controls()
        self._update_action_state()

    def _scan_qa_reports(self) -> None:
        report_paths = self.qa_reports_zone.file_paths()
        self.qa_sheet_configs = []
        self.qa_scan_warnings = []
        if not report_paths:
            self._update_qa_mapping_status_label()
            self._update_qa_controls()
            return

        verifier = QAVerifier()
        scan_result: QAScanResult = verifier.scan_reports(report_paths)
        self.qa_sheet_configs = scan_result.sheet_configs
        self.qa_scan_warnings = scan_result.warnings
        self._update_qa_mapping_status_label()
        self._update_qa_controls()

    def _open_qa_mapping_dialog(self) -> None:
        if not self.qa_sheet_configs:
            QMessageBox.information(self, "QA Verify", "Load QA reports first.")
            return
        dialog = QAColumnMappingDialog(self.qa_sheet_configs, self)
        if dialog.exec():
            self.qa_sheet_configs = dialog.sheet_configs()
            self._update_qa_mapping_status_label()
            self._update_qa_controls()
            self._update_action_state()

    def _update_qa_mapping_status_label(self) -> None:
        if not self.qa_sheet_configs:
            self.qa_mapping_status_label.setText("Load QA reports to detect columns.")
            return
        complete = sum(1 for item in self.qa_sheet_configs if item.mapping.is_complete())
        total = len(self.qa_sheet_configs)
        text = f"Sheets: {total}, mapped: {complete}, unresolved: {total - complete}"
        if self.qa_scan_warnings:
            text += f" (warnings: {len(self.qa_scan_warnings)})"
        self.qa_mapping_status_label.setText(text)

    def _qa_can_run(self) -> bool:
        if not self.qa_sheet_configs:
            return False
        if not self.qa_final_zone.file_paths():
            return False
        return any(item.mapping.is_complete() for item in self.qa_sheet_configs)

    def _update_qa_controls(self) -> None:
        has_sheets = bool(self.qa_sheet_configs)
        self.qa_map_columns_btn.setEnabled(has_sheets)
        self.qa_export_btn.setEnabled(self.qa_result is not None and bool(self.qa_result.rows))

    def _populate_qa_results_table(self, result: QAVerificationResult | None) -> None:
        if result is None:
            for value_label in self.qa_summary_labels.values():
                value_label.setText("0")
            self._update_qa_controls()
            return

        self._update_qa_summary(result)
        self._update_qa_controls()

    def _update_qa_summary(self, result: QAVerificationResult) -> None:
        for status, label in self.qa_summary_labels.items():
            label.setText(str(result.status_counts.get(status, 0)))

    def _export_qa_results(self) -> None:
        if self.qa_result is None or not self.qa_result.rows:
            QMessageBox.information(self, "QA Verify", "No QA verification results to export.")
            return

        output_dir = self.output_line.text().strip() or "./output/"
        default_name = (
            f"qa_verify_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        default_path = str(Path(output_dir) / default_name)
        selected_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export QA Verification",
            default_path,
            "Excel file (*.xlsx)",
        )
        if not selected_path:
            return
        try:
            exported = QAVerifier().export_to_excel(self.qa_result, selected_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export failed", str(exc))
            return
        self.statusBar().showMessage(f"QA results exported: {exported}")

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
        payload: dict[str, object]
        if self.current_mode == self.MODE_FILE:
            output_dir = self.output_line.text().strip()
            if not output_dir:
                return
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
            excel_source_col_a: str | None = None
            excel_source_col_b: str | None = None
            has_excel_pairs = any(
                Path(file_a).suffix.lower() in {".xlsx", ".xls"}
                for file_a, _ in pairs
            )
            if has_excel_pairs:
                try:
                    excel_source_col_a = self._normalize_excel_column_input(
                        self.excel_source_col_a_input.text()
                    )
                    excel_source_col_b = self._normalize_excel_column_input(
                        self.excel_source_col_b_input.text()
                    )
                except ValueError as exc:
                    QMessageBox.warning(self, "Invalid Excel source column", str(exc))
                    return
            payload = {
                "pairs": pairs,
                "output_dir": output_dir,
                "excel_source_col_a": excel_source_col_a,
                "excel_source_col_b": excel_source_col_b,
            }
        elif self.current_mode == self.MODE_VERSIONS:
            output_dir = self.output_line.text().strip()
            if not output_dir:
                return
            files = [self.version_list.item(i).text() for i in range(self.version_list.count())]
            if len(files) < 2:
                return
            payload = {"files": files, "output_dir": output_dir}
        else:
            if not self._qa_can_run():
                QMessageBox.warning(
                    self,
                    "QA Verify",
                    "Load QA reports, map Source/Original/QA mark for at least one sheet, "
                    "and add final XLIFF files.",
                )
                return
            payload = {
                "sheet_configs": [item.to_dict() for item in self.qa_sheet_configs],
                "final_files": self.qa_final_zone.file_paths(),
            }

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
                outputs = [str(path) for path in payload.get("outputs", [])]
                self.last_html_report = next(
                    (path for path in outputs if path.lower().endswith(".html")), None
                )
                self.last_excel_report = next(
                    (path for path in outputs if path.lower().endswith(".xlsx")), None
                )
                statistics = payload.get("statistics")
                changed_total = self._changed_count(statistics)
                self.statusBar().showMessage(
                    "Done: pairs={total}, compared={ok}, errors={err}, changed={changed}".format(
                        total=len(file_results),
                        ok=len(successful),
                        err=len(failed),
                        changed=changed_total,
                    )
                )
                if (
                    self.last_html_report
                    and statistics is not None
                    and not self._statistics_has_changes(statistics)
                ):
                    QMessageBox.information(
                        self,
                        "No changes",
                        "Правок не найдено: отчет пустой.",
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
                    if self.last_html_report and not self._statistics_has_changes(stats):
                        QMessageBox.information(
                            self,
                            "No changes",
                            "Правок не найдено: отчет пустой.",
                        )
        elif mode == self.MODE_VERSIONS:
            result = payload["result"]
            self.last_html_report = result.summary_report_path
            self.last_excel_report = None
            self.statusBar().showMessage(
                f"Done: {len(result.comparisons)} comparisons generated"
            )
        elif mode == self.MODE_QA_VERIFY:
            result = payload["result"]
            self.qa_result = result
            self._populate_qa_results_table(result)
            self.statusBar().showMessage(
                "Done: rows={rows}, applied={applied}, not_applied={not_applied}, "
                "cannot_verify={cannot_verify}, not_applicable={na}".format(
                    rows=result.total_rows,
                    applied=result.status_counts.get(STATUS_APPLIED, 0),
                    not_applied=result.status_counts.get(STATUS_NOT_APPLIED, 0),
                    cannot_verify=result.status_counts.get(STATUS_CANNOT_VERIFY, 0),
                    na=result.status_counts.get(STATUS_NOT_APPLICABLE, 0),
                )
            )
            if result.warnings:
                preview = "\n".join(f"- {item}" for item in result.warnings[:8])
                QMessageBox.warning(self, "QA Verify warnings", preview)

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
        elif self.current_mode == self.MODE_QA_VERIFY:
            enabled = bool(self._qa_can_run())
        self.compare_btn.setEnabled(enabled and self.worker is None)
        self._update_qa_controls()

    @staticmethod
    def _changed_count(statistics: object) -> int:
        if statistics is None:
            return 0
        if isinstance(statistics, dict):
            return int(
                statistics.get("added", 0)
                + statistics.get("deleted", 0)
                + statistics.get("modified", 0)
                + statistics.get("moved", 0)
            )
        return int(
            getattr(statistics, "added", 0)
            + getattr(statistics, "deleted", 0)
            + getattr(statistics, "modified", 0)
            + getattr(statistics, "moved", 0)
        )

    @staticmethod
    def _statistics_has_changes(statistics: object) -> bool:
        return MainWindow._changed_count(statistics) > 0

    def _supported_filter(self) -> str:
        if not self.supported_extensions:
            return "All files (*.*)"
        patterns = " ".join(f"*{ext}" for ext in self.supported_extensions)
        return f"Supported files ({patterns});;All files (*.*)"

    @staticmethod
    def _normalize_excel_column_input(value: str) -> str | None:
        text = value.strip()
        if not text:
            return None
        upper = text.upper()
        if upper.isdigit():
            number = int(upper)
            if number <= 0:
                raise ValueError(
                    "Excel source column must be a positive number or letters (A, B, ...)."
                )
            return str(number)
        if upper.isalpha():
            return upper
        raise ValueError(
            "Excel source column must contain only letters (A, B, ...) "
            "or a positive column number."
        )

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
QFrame#qaSummaryBox {
  background: #ffffff;
  border: 1px solid #dbe4f0;
  border-radius: 8px;
}
QLabel#qaSummaryValue {
  font-weight: 700;
  color: #0f172a;
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
