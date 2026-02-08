from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QThread, pyqtSignal

from core.orchestrator import Orchestrator
from core.qa_verify import QASheetConfig, QAVerifier


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
                if "pairs" in self.payload:
                    pairs = list(self.payload["pairs"])
                    file_results: list[dict[str, Any]] = []
                    total = max(1, len(pairs))
                    for index, pair in enumerate(pairs, start=1):
                        file_a, file_b = pair
                        pair_name = self._pair_folder_name(index, str(file_a), str(file_b))
                        pair_output = str(Path(self.payload["output_dir"]) / pair_name)
                        self._emit_progress(
                            f"Comparing {index}/{len(pairs)}: {Path(file_a).name} vs {Path(file_b).name}",
                            (index - 1) / total,
                        )
                        try:
                            outputs = orchestrator.compare_files(
                                str(file_a),
                                str(file_b),
                                pair_output,
                            )
                            file_results.append(
                                {
                                    "file_a": str(file_a),
                                    "file_b": str(file_b),
                                    "outputs": outputs,
                                    "comparison": orchestrator.last_result,
                                    "error": None,
                                }
                            )
                        except Exception as exc:  # pragma: no cover - signal path
                            file_results.append(
                                {
                                    "file_a": str(file_a),
                                    "file_b": str(file_b),
                                    "outputs": [],
                                    "comparison": None,
                                    "error": str(exc),
                                }
                            )

                    self._emit_progress("Done", 1.0)
                    self.finished.emit(
                        {
                            "mode": "file",
                            "multi": True,
                            "file_results": file_results,
                        }
                    )
                    return

                outputs = orchestrator.compare_files(
                    self.payload["file_a"],
                    self.payload["file_b"],
                    self.payload["output_dir"],
                )
                self.finished.emit(
                    {
                        "mode": "file",
                        "multi": False,
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

            if self.mode == "qa_verify":
                sheet_configs = [
                    QASheetConfig.from_dict(item)
                    if isinstance(item, dict)
                    else item
                    for item in self.payload.get("sheet_configs", [])
                ]
                final_files = [str(item) for item in self.payload.get("final_files", [])]
                verifier = QAVerifier(on_progress=self._emit_progress)
                result = verifier.verify(sheet_configs, final_files)
                self.finished.emit({"mode": "qa_verify", "result": result})
                return

            self.error.emit(f"Unknown worker mode: {self.mode}")
        except Exception as exc:  # pragma: no cover - signal path
            self.error.emit(str(exc))

    def _emit_progress(self, message: str, value: float) -> None:
        self.progress.emit(message, value)

    @staticmethod
    def _pair_folder_name(index: int, file_a: str, file_b: str) -> str:
        stem_a = Path(file_a).stem
        stem_b = Path(file_b).stem
        digest = hashlib.sha1(
            f"{stem_a}|{stem_b}".encode("utf-8", errors="ignore")
        ).hexdigest()[:10]
        part_a = ComparisonWorker._safe_name_part(stem_a)
        part_b = ComparisonWorker._safe_name_part(stem_b)
        return f"{index:03d}_{part_a}_vs_{part_b}_{digest}"

    @staticmethod
    def _safe_name_part(value: str, max_len: int = 36) -> str:
        safe = value.replace(" ", "_")
        for bad in '<>:"/\\|?*':
            safe = safe.replace(bad, "_")
        safe = safe.strip("._")
        if not safe:
            safe = "file"
        if len(safe) > max_len:
            safe = safe[:max_len].rstrip("._")
        return safe or "file"
