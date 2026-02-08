from __future__ import annotations

from typing import Any

from PyQt6.QtCore import QThread, pyqtSignal

from core.orchestrator import Orchestrator


class ComparisonWorker(QThread):
    progress = pyqtSignal(str, float)
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, mode: str, payload: dict[str, Any], parent=None) -> None:
        super().__init__(parent)
        self.mode = mode
        self.payload = payload

    def run(self) -> None:
        orchestrator = Orchestrator(on_progress=self._emit_progress)
        try:
            if self.mode == "file":
                outputs = orchestrator.compare_files(
                    self.payload["file_a"],
                    self.payload["file_b"],
                    self.payload["output_dir"],
                )
                self.finished.emit(
                    {
                        "mode": "file",
                        "outputs": outputs,
                        "comparison": orchestrator.last_result,
                    }
                )
                return

            if self.mode == "batch":
                result = orchestrator.compare_folders(
                    self.payload["folder_a"],
                    self.payload["folder_b"],
                    self.payload["output_dir"],
                )
                self.finished.emit({"mode": "batch", "result": result})
                return

            if self.mode == "versions":
                result = orchestrator.compare_versions(
                    self.payload["files"],
                    self.payload["output_dir"],
                )
                self.finished.emit({"mode": "versions", "result": result})
                return

            self.error.emit(f"Unknown worker mode: {self.mode}")
        except Exception as exc:  # pragma: no cover - signal path
            self.error.emit(str(exc))

    def _emit_progress(self, message: str, value: float) -> None:
        self.progress.emit(message, value)

