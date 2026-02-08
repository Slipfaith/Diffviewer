from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PyQt6.QtCore import QEvent, QMimeData, QSize, Qt, QUrl, pyqtSignal
from PyQt6.QtGui import QDrag, QDragEnterEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class TileVisualState:
    matched: bool = False
    unmatched: bool = False
    selected: bool = False
    candidate: bool = False


class _FileTileWidget(QFrame):
    def __init__(self, filename: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._filename = filename
        self._label = QLabel(filename, self)
        self._label.setWordWrap(False)
        self._label.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        self._label.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.addWidget(self._label)
        self._refresh_text()
        self.apply_visual(TileVisualState())

    def set_filename(self, filename: str) -> None:
        self._filename = filename
        self._refresh_text()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._refresh_text()

    def _refresh_text(self) -> None:
        available_width = max(24, self._label.width() - 2)
        text = self._label.fontMetrics().elidedText(
            self._filename,
            Qt.TextElideMode.ElideRight,
            available_width,
        )
        tooltip = self._wrap_tooltip(self._filename)
        self._label.setText(text)
        self._label.setToolTip(tooltip)
        self.setToolTip(tooltip)

    @staticmethod
    def _wrap_tooltip(text: str, line_len: int = 32) -> str:
        if len(text) <= line_len:
            return text

        chunks: list[str] = []
        remaining = text
        while len(remaining) > line_len:
            split_at = max(
                remaining.rfind(" ", 0, line_len + 1),
                remaining.rfind("_", 0, line_len + 1),
                remaining.rfind("-", 0, line_len + 1),
                remaining.rfind(".", 0, line_len + 1),
            )
            if split_at <= 0:
                split_at = line_len
            chunks.append(remaining[:split_at].rstrip(" _-."))
            remaining = remaining[split_at:].lstrip(" _-.")
        if remaining:
            chunks.append(remaining)
        return "\n".join(chunks) if chunks else text

    def apply_visual(self, state: TileVisualState) -> None:
        background = "#ffffff"
        border = "#cbd5e1"
        text_color = "#1f2937"
        if state.matched:
            background = "#ecfdf3"
            border = "#86efac"
        elif state.unmatched:
            background = "#fef2f2"
            border = "#fecaca"
        if state.selected:
            background = "#fff7ed"
            border = "#fdba74"

        border_style = "dashed" if state.candidate else "solid"
        self.setStyleSheet(
            f"""
QFrame {{
  background: {background};
  border: 1px {border_style} {border};
  border-radius: 8px;
}}
QLabel {{
  color: {text_color};
  font-size: 12px;
  font-weight: 500;
  background: transparent;
  border: none;
}}
"""
        )


class _FileTileListWidget(QListWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        self.setDefaultDropAction(Qt.DropAction.CopyAction)

    def startDrag(self, supported_actions) -> None:  # type: ignore[override]
        selected = self.selectedItems()
        if not selected:
            return
        urls: list[QUrl] = []
        lines: list[str] = []
        for item in selected:
            path = item.data(Qt.ItemDataRole.UserRole)
            if not isinstance(path, str):
                continue
            lines.append(path)
            urls.append(QUrl.fromLocalFile(path))
        if not urls:
            return
        mime = QMimeData()
        mime.setUrls(urls)
        mime.setText("\n".join(lines))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.CopyAction)


class FileTileDropZone(QFrame):
    files_changed = pyqtSignal(list)
    file_left_clicked = pyqtSignal(str)

    def __init__(
        self,
        title: str,
        *,
        allowed_extensions: Iterable[str] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.title = title
        self.allowed_extensions = {
            ext.lower() for ext in (allowed_extensions or []) if ext
        }

        self.setObjectName("fileTileDropZone")
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.Shape.StyledPanel)

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        self.title_label = QLabel(title, self)
        self.title_label.setObjectName("fileTileDropZoneTitle")
        root.addWidget(self.title_label)

        self.list_widget = _FileTileListWidget(self)
        self.list_widget.setObjectName("fileTileList")
        self.list_widget.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        self.list_widget.setAlternatingRowColors(False)
        self.list_widget.setSpacing(6)
        self.list_widget.setMinimumHeight(200)
        self.list_widget.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        self.list_widget.customContextMenuRequested.connect(self._show_context_menu)
        self.list_widget.installEventFilter(self)
        self.list_widget.viewport().installEventFilter(self)
        root.addWidget(self.list_widget, 1)

        self.hint_label = QLabel(
            "Drop files here\nor double-click to browse",
            self.list_widget.viewport(),
        )
        self.hint_label.setObjectName("fileTileHint")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True,
        )
        self._update_hint()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self.hint_label.resize(self.list_widget.viewport().size())

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._extract_valid_paths(event):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if self._extract_valid_paths(event):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = self._extract_valid_paths(event)
        if not paths:
            event.ignore()
            return
        self.add_files(paths)
        event.acceptProposedAction()

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.open_file_dialog()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def eventFilter(self, watched, event) -> bool:  # type: ignore[override]
        if watched in {self.list_widget, self.list_widget.viewport()}:
            if event.type() == QEvent.Type.MouseButtonDblClick:
                mouse_event = event
                if mouse_event.button() == Qt.MouseButton.LeftButton:
                    self.open_file_dialog()
                    return True
        return super().eventFilter(watched, event)

    def add_files(self, paths: Iterable[str]) -> None:
        existing = set(self.file_paths())
        added = False
        for raw_path in paths:
            normalized = self._normalize_path(raw_path)
            if normalized is None:
                continue
            candidate = Path(normalized)
            files_to_add = (
                self._files_from_directory(candidate)
                if candidate.is_dir()
                else [candidate]
            )
            for file_path in files_to_add:
                file_path_str = str(file_path)
                if file_path_str in existing:
                    continue
                item = QListWidgetItem()
                item.setSizeHint(self._tile_size_hint())
                item.setData(Qt.ItemDataRole.DisplayRole, file_path.name)
                item.setData(Qt.ItemDataRole.UserRole, file_path_str)
                item.setToolTip(_FileTileWidget._wrap_tooltip(file_path.name))
                widget = _FileTileWidget(file_path.name, self.list_widget)
                self.list_widget.addItem(item)
                self.list_widget.setItemWidget(item, widget)
                existing.add(file_path_str)
                added = True

        if added:
            self._emit_changed()

    def remove_file(self, file_path: str) -> None:
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.data(Qt.ItemDataRole.UserRole) == file_path:
                self.list_widget.takeItem(row)
                self._emit_changed()
                return

    def clear_files(self) -> None:
        if self.list_widget.count() == 0:
            return
        self.list_widget.clear()
        self._emit_changed()

    def file_paths(self) -> list[str]:
        paths: list[str] = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            path = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(path, str):
                paths.append(path)
        return paths

    def apply_states(self, states: dict[str, TileVisualState]) -> None:
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            path = item.data(Qt.ItemDataRole.UserRole)
            widget = self.list_widget.itemWidget(item)
            if not isinstance(path, str) or not isinstance(widget, _FileTileWidget):
                continue
            widget.apply_visual(states.get(path, TileVisualState()))

    def open_file_dialog(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            f"Select files: {self.title}",
            "",
            self._file_filter(),
        )
        if selected:
            self.add_files(selected)

    def _show_context_menu(self, pos) -> None:
        item = self.list_widget.itemAt(pos)
        if item is None:
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(path, str):
            return

        menu = QMenu(self.list_widget)
        remove_action = menu.addAction("Remove from list")
        action = menu.exec(self.list_widget.viewport().mapToGlobal(pos))
        if action == remove_action:
            self.remove_file(path)

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        path = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(path, str):
            self.file_left_clicked.emit(path)

    def _emit_changed(self) -> None:
        self._update_hint()
        self.files_changed.emit(self.file_paths())

    def _update_hint(self) -> None:
        self.hint_label.setVisible(self.list_widget.count() == 0)

    def _extract_valid_paths(self, event) -> list[str]:
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            return []

        paths: list[str] = []
        for url in mime_data.urls():
            local = url.toLocalFile()
            normalized = self._normalize_path(local)
            if normalized is not None:
                paths.append(normalized)
        return paths

    def _normalize_path(self, path: str) -> str | None:
        if not path:
            return None
        candidate = Path(path)
        if not candidate.exists():
            return None
        if candidate.is_file():
            if self.allowed_extensions and candidate.suffix.lower() not in self.allowed_extensions:
                return None
        elif not candidate.is_dir():
            return None
        try:
            return str(candidate.resolve())
        except Exception:
            return str(candidate)

    def _files_from_directory(self, directory: Path) -> list[Path]:
        files: list[Path] = []
        try:
            for path in sorted(directory.iterdir()):
                if not path.is_file():
                    continue
                if self.allowed_extensions and path.suffix.lower() not in self.allowed_extensions:
                    continue
                files.append(path.resolve())
        except Exception:
            return []
        return files

    @staticmethod
    def _tile_size_hint():
        probe = QLabel("x")
        font_metrics = probe.fontMetrics()
        line_height = font_metrics.height()
        width = max(220, line_height * 12)
        height = max(42, line_height * 2)
        return QSize(width, height)

    def _file_filter(self) -> str:
        if not self.allowed_extensions:
            return "All files (*.*)"
        patterns = " ".join(f"*{ext}" for ext in sorted(self.allowed_extensions))
        return f"Supported files ({patterns});;All files (*.*)"
