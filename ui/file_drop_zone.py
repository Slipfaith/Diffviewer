from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import QLabel, QVBoxLayout, QWidget


class FileDropZone(QWidget):
    file_dropped = pyqtSignal(str)

    def __init__(
        self,
        title: str,
        *,
        accept_directories: bool = False,
        allowed_extensions: Iterable[str] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.accept_directories = accept_directories
        self.allowed_extensions = {
            ext.lower() for ext in (allowed_extensions or []) if ext
        }
        self._drag_active = False

        self.setAcceptDrops(True)
        self.setMinimumHeight(108)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("dropZoneTitle")
        layout.addWidget(self.title_label)

        self.path_label = QLabel("Drop here or use Browse")
        self.path_label.setObjectName("dropZonePath")
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)
        layout.addStretch(1)

        self._apply_style()

    def set_path(self, path: str) -> None:
        self.path_label.setText(path)
        self.path_label.setToolTip(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        path = self._first_valid_path(event)
        if path is None:
            event.ignore()
            return
        self._drag_active = True
        self._apply_style()
        event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:  # type: ignore[override]
        self._drag_active = False
        self._apply_style()
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        path = self._first_valid_path(event)
        self._drag_active = False
        self._apply_style()
        if path is None:
            event.ignore()
            return
        self.set_path(path)
        self.file_dropped.emit(path)
        event.acceptProposedAction()

    def _first_valid_path(self, event: QDragEnterEvent | QDropEvent) -> str | None:
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            return None

        for url in mime_data.urls():
            local_path = url.toLocalFile()
            if not local_path:
                continue
            if self._is_supported(local_path):
                return local_path
        return None

    def _is_supported(self, path: str) -> bool:
        candidate = Path(path)
        if self.accept_directories:
            return candidate.is_dir()
        if candidate.is_dir():
            return False
        if not self.allowed_extensions:
            return True
        return candidate.suffix.lower() in self.allowed_extensions

    def _apply_style(self) -> None:
        border_color = "#1d4ed8" if self._drag_active else "#94a3b8"
        bg_color = "#eff6ff" if self._drag_active else "#ffffff"
        self.setStyleSheet(
            f"""
QWidget {{
  background: {bg_color};
  border: 2px dashed {border_color};
  border-radius: 10px;
}}
QLabel#dropZoneTitle {{
  color: #334155;
  font-size: 12px;
  font-weight: 600;
  border: none;
  background: transparent;
}}
QLabel#dropZonePath {{
  color: #475569;
  font-size: 12px;
  border: none;
  background: transparent;
}}
"""
        )

