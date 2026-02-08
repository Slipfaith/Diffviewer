from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import logging
from pathlib import Path
from typing import Callable

from core.diff_engine import DiffEngine
from core.models import (
    BatchFileResult,
    BatchResult,
    ComparisonResult,
    MultiVersionResult,
    ParseError,
    UnsupportedFormatError,
)
from core.registry import ParserRegistry, ReporterRegistry
from reporters.docx_reporter import DocxTrackChangesReporter
from reporters.excel_reporter import ExcelReporter
from reporters.html_reporter import HtmlReporter
from reporters.summary_reporter import SummaryReporter


logger = logging.getLogger(__name__)


@dataclass
class Orchestrator:
    on_progress: Callable[[str, float], None] | None = None
    last_result: ComparisonResult | None = None

    def __post_init__(self) -> None:
        ParserRegistry.discover()
        ReporterRegistry.discover()

    def compare_files(self, file_a: str, file_b: str, output_dir: str) -> list[str]:
        path_a = Path(file_a)
        path_b = Path(file_b)
        ext_a = path_a.suffix.lower()
        ext_b = path_b.suffix.lower()
        if ext_a != ext_b:
            raise UnsupportedFormatError(f"{ext_a} vs {ext_b}")

        self._progress("Selecting parser", 0.1)
        try:
            parser = ParserRegistry.get_parser(str(path_a))
        except UnsupportedFormatError as exc:
            raise UnsupportedFormatError(ext_a) from exc

        self._progress("Parsing file A", 0.2)
        try:
            doc_a = parser.parse(str(path_a))
        except ParseError as exc:
            raise ParseError(file_a, exc.reason) from exc

        self._progress("Parsing file B", 0.4)
        try:
            doc_b = parser.parse(str(path_b))
        except ParseError as exc:
            raise ParseError(file_b, exc.reason) from exc

        self._progress("Comparing documents", 0.6)
        result = DiffEngine.compare(doc_a, doc_b)
        self.last_result = result

        self._progress("Generating report", 0.8)
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        timestamp_label = datetime.now().strftime("%d-%m-%y--%H-%M")
        base_name = f"changereport_{timestamp_label}"
        html_path = output_dir_path / f"{base_name}.html"
        excel_path = output_dir_path / f"{base_name}.xlsx"

        outputs: list[str] = []
        if ext_a == ".docx":
            docx_reporter = DocxTrackChangesReporter()
            if docx_reporter.is_available():
                output_path = output_dir_path / f"{base_name}{docx_reporter.output_extension}"
                outputs.append(docx_reporter.generate(result, str(output_path)))
                outputs.append(HtmlReporter().generate(result, str(html_path)))
                outputs.append(ExcelReporter().generate(result, str(excel_path)))
            else:
                logger.warning(
                    "Microsoft Word not found, generating HTML+Excel reports only"
                )
                outputs.append(HtmlReporter().generate(result, str(html_path)))
                outputs.append(ExcelReporter().generate(result, str(excel_path)))
        else:
            outputs.append(HtmlReporter().generate(result, str(html_path)))
            outputs.append(ExcelReporter().generate(result, str(excel_path)))

        self._progress("Done", 1.0)
        return outputs

    def compare_folders(self, folder_a: str, folder_b: str, output_dir: str) -> BatchResult:
        path_a = Path(folder_a)
        path_b = Path(folder_b)
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        files_a = {p.name.lower(): p for p in path_a.iterdir() if p.is_file()}
        files_b = {p.name.lower(): p for p in path_b.iterdir() if p.is_file()}

        all_keys = sorted(set(files_a.keys()) | set(files_b.keys()))
        results: list[BatchFileResult] = []

        total = len(all_keys) if all_keys else 1
        for index, key in enumerate(all_keys, start=1):
            self._progress(f"Comparing file {index}/{len(all_keys)}...", index / total)
            file_a = files_a.get(key)
            file_b = files_b.get(key)

            if file_a is None and file_b is not None:
                results.append(
                    BatchFileResult(
                        filename=file_b.name,
                        status="only_in_b",
                    )
                )
                continue
            if file_b is None and file_a is not None:
                results.append(
                    BatchFileResult(
                        filename=file_a.name,
                        status="only_in_a",
                    )
                )
                continue

            if file_a is None or file_b is None:
                continue

            try:
                pair_output_dir = output_dir_path / self._safe_stem(file_a.name)
                outputs = self.compare_files(
                    str(file_a),
                    str(file_b),
                    str(pair_output_dir),
                )
                stats = self.last_result.statistics if self.last_result is not None else None
                results.append(
                    BatchFileResult(
                        filename=file_a.name,
                        status="compared",
                        report_paths=outputs,
                        statistics=stats,
                        comparison=self.last_result,
                    )
                )
            except Exception as exc:
                results.append(
                    BatchFileResult(
                        filename=file_a.name,
                        status="error",
                        error_message=str(exc),
                    )
                )

        batch_result = BatchResult(
            folder_a=str(path_a),
            folder_b=str(path_b),
            files=results,
        )

        summary_reporter = SummaryReporter()
        summary_path = output_dir_path / "batch_summary.html"
        summary_excel_path = output_dir_path / "batch_summary.xlsx"
        batch_result.summary_report_path = summary_reporter.generate_batch(
            batch_result,
            str(summary_path),
        )
        batch_result.summary_excel_path = summary_reporter.generate_batch_excel(
            batch_result,
            str(summary_excel_path),
        )
        return batch_result

    def compare_versions(self, files: list[str], output_dir: str) -> MultiVersionResult:
        if len(files) < 2:
            return MultiVersionResult(file_paths=files, comparisons=[], report_paths=[])

        extensions = {Path(path).suffix.lower() for path in files}
        if len(extensions) != 1:
            raise UnsupportedFormatError("mixed formats")

        comparisons: list[ComparisonResult] = []
        report_paths: list[list[str]] = []
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        for idx in range(len(files) - 1):
            self._progress(
                f"Comparing version {idx + 1}/{len(files) - 1}...",
                (idx + 1) / (len(files) - 1),
            )
            pair_output_dir = output_dir_path / f"v{idx + 1}_to_v{idx + 2}"
            outputs = self.compare_files(
                files[idx],
                files[idx + 1],
                str(pair_output_dir),
            )
            if self.last_result is not None:
                comparisons.append(self.last_result)
            report_paths.append(outputs)

        multi = MultiVersionResult(
            file_paths=files,
            comparisons=comparisons,
            report_paths=report_paths,
        )

        summary_path = output_dir_path / "versions_summary.html"
        SummaryReporter().generate_versions(multi, str(summary_path))
        return multi

    def _progress(self, message: str, value: float) -> None:
        if self.on_progress is not None:
            self.on_progress(message, value)

    @staticmethod
    def _safe_stem(filename: str) -> str:
        safe = filename.replace(" ", "_")
        for bad in '<>:"/\\|?*':
            safe = safe.replace(bad, "_")
        return safe
