from __future__ import annotations

import logging
import subprocess
from pathlib import Path

from core.models import ComparisonResult
from reporters.base import BaseReporter
from reporters.excel_reporter import ExcelReporter
from reporters.html_reporter import HtmlReporter


logger = logging.getLogger(__name__)


class DocxTrackChangesReporter(BaseReporter):
    name = "DOCX Track Changes Reporter"
    output_extension = ".docx"
    supports_rich_text = True

    def __init__(self, author: str = "Change Tracker") -> None:
        self.author = author
        root = Path(__file__).resolve().parents[1]
        self._exe_path = root / "tools" / "docx_compare.exe"

    def is_available(self) -> bool:
        return self._exe_path.exists()

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        if not self.is_available():
            logger.warning(
                "DOCX Track Changes module not found, falling back to HTML+Excel report"
            )
            html_path = output_file.with_suffix(".html")
            xlsx_path = output_file.with_suffix(".xlsx")
            HtmlReporter().generate(result, str(html_path))
            ExcelReporter().generate(result, str(xlsx_path))
            return str(html_path)

        output_file.parent.mkdir(parents=True, exist_ok=True)
        cmd = [
            str(self._exe_path),
            result.file_a.file_path,
            result.file_b.file_path,
            str(output_file),
            "--author",
            self.author,
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True)
        if completed.returncode != 0:
            logger.error("docx_compare.exe failed: %s", completed.stderr.strip())
            raise RuntimeError(completed.stderr.strip() or "docx_compare.exe failed")

        return str(output_file)
