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
                    if len(pairs) <= 1:
                        if not pairs:
                            self.finished.emit(
                                {
                                    "mode": "file",
                                    "multi": True,
                                    "outputs": [],
                                    "file_results": [],
                                    "statistics": None,
                                }
                            )
                            return
                        file_a, file_b = pairs[0]
                        outputs = orchestrator.compare_files(
                            str(file_a),
                            str(file_b),
                            str(self.payload["output_dir"]),
                            excel_source_column_a=self.payload.get("excel_source_col_a"),
                            excel_source_column_b=self.payload.get("excel_source_col_b"),
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

                    pair_result = orchestrator.compare_file_pairs(
                        pairs=[(str(file_a), str(file_b)) for file_a, file_b in pairs],
                        output_dir=str(self.payload["output_dir"]),
                        excel_source_column_a=self.payload.get("excel_source_col_a"),
                        excel_source_column_b=self.payload.get("excel_source_col_b"),
                    )
                    self.finished.emit(
                        {
                            "mode": "file",
                            "multi": True,
                            "outputs": pair_result.get("outputs", []),
                            "file_results": pair_result.get("file_results", []),
                            "statistics": pair_result.get("statistics"),
                        }
                    )
                    return

                outputs = orchestrator.compare_files(
                    self.payload["file_a"],
                    self.payload["file_b"],
                    self.payload["output_dir"],
                    excel_source_column_a=self.payload.get("excel_source_col_a"),
                    excel_source_column_b=self.payload.get("excel_source_col_b"),
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
