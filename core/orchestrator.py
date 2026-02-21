from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import logging
from pathlib import Path
from typing import Any, Callable

from core.diff_engine import DiffEngine
from core.models import (
    BatchFileResult,
    BatchResult,
    ChangeStatistics,
    ComparisonResult,
    MultiVersionResult,
    OneVsAllResult,
    ParseError,
    ParsedDocument,
    UnsupportedFormatError,
)

_ONE_VS_ALL_EXTENSIONS = {".xliff", ".xlf", ".sdlxliff", ".mqxliff"}
from core.registry import ParserRegistry, ReporterRegistry
from parsers.base import BaseParser
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

    def compare_files(
        self,
        file_a: str,
        file_b: str,
        output_dir: str,
        *,
        excel_source_column_a: str | int | None = None,
        excel_source_column_b: str | int | None = None,
    ) -> list[str]:
        path_a = Path(file_a)
        path_b = Path(file_b)
        ext_a = path_a.suffix.lower()
        ext_b = path_b.suffix.lower()
        if ext_a != ext_b:
            raise UnsupportedFormatError(f"{ext_a} vs {ext_b}")

        self._progress("Selecting parser", 0.1)
        try:
            parser_a = ParserRegistry.get_parser(str(path_a))
            parser_b = ParserRegistry.get_parser(str(path_b))
        except UnsupportedFormatError as exc:
            raise UnsupportedFormatError(ext_a) from exc

        self._configure_excel_source_column(
            parser=parser_a,
            extension=ext_a,
            source_column=excel_source_column_a,
            file_path=file_a,
        )
        self._configure_excel_source_column(
            parser=parser_b,
            extension=ext_b,
            source_column=excel_source_column_b,
            file_path=file_b,
        )

        self._progress("Parsing file A", 0.2)
        decode_entities = self._should_decode_entities(ext_a)
        try:
            doc_a = parser_a.parse(str(path_a))
        except ParseError as exc:
            raise ParseError(file_a, exc.reason) from exc
        self._normalize_document_text_entities(doc_a, decode_entities=decode_entities)

        self._progress("Parsing file B", 0.4)
        try:
            doc_b = parser_b.parse(str(path_b))
        except ParseError as exc:
            raise ParseError(file_b, exc.reason) from exc
        self._normalize_document_text_entities(doc_b, decode_entities=decode_entities)

        self._progress("Comparing documents", 0.6)
        result = DiffEngine.compare(doc_a, doc_b)
        self.last_result = result

        self._progress("Generating report", 0.8)
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        timestamp_label = datetime.now().strftime("%d-%m-%y--%H-%M-%S")
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

    def compare_file_pairs(
        self,
        pairs: list[tuple[str, str]],
        output_dir: str,
        *,
        excel_source_column_a: str | int | None = None,
        excel_source_column_b: str | int | None = None,
    ) -> dict[str, object]:
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        successful_results: list[tuple[str, ComparisonResult]] = []
        file_results: list[dict[str, object]] = []
        total = len(pairs) if pairs else 1

        for index, (file_a, file_b) in enumerate(pairs, start=1):
            self._progress(
                f"Comparing {index}/{len(pairs)}: {Path(file_a).name} vs {Path(file_b).name}",
                (index - 1) / total,
            )
            try:
                result = self._compare_pair_without_reports(
                    file_a,
                    file_b,
                    excel_source_column_a=excel_source_column_a,
                    excel_source_column_b=excel_source_column_b,
                )
                self.last_result = result
                pair_label = f"{Path(file_a).name} vs {Path(file_b).name}"
                successful_results.append((pair_label, result))
                file_results.append(
                    {
                        "file_a": file_a,
                        "file_b": file_b,
                        "comparison": result,
                        "error": None,
                    }
                )
            except Exception as exc:
                file_results.append(
                    {
                        "file_a": file_a,
                        "file_b": file_b,
                        "comparison": None,
                        "error": str(exc),
                    }
                )

        outputs: list[str] = []
        aggregate_statistics = ChangeStatistics.from_changes([])
        if successful_results:
            self._progress("Generating consolidated report", 0.85)
            timestamp_label = datetime.now().strftime("%d-%m-%y--%H-%M-%S")
            base_name = f"changereport_multi_{timestamp_label}"
            html_path = output_dir_path / f"{base_name}.html"
            excel_path = output_dir_path / f"{base_name}.xlsx"
            outputs.append(HtmlReporter().generate_multi(successful_results, str(html_path)))
            outputs.append(ExcelReporter().generate_multi(successful_results, str(excel_path)))

            all_changes = [
                change
                for _, comparison in successful_results
                for change in comparison.changes
            ]
            aggregate_statistics = ChangeStatistics.from_changes(all_changes)

        self._progress("Done", 1.0)
        return {
            "outputs": outputs,
            "file_results": file_results,
            "statistics": aggregate_statistics,
        }

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

        path_objects = [Path(path) for path in files]
        base_path = path_objects[0]

        self._progress("Selecting parser", 0.05)
        try:
            parser = ParserRegistry.get_parser(str(base_path))
        except UnsupportedFormatError as exc:
            raise UnsupportedFormatError(base_path.suffix.lower()) from exc

        documents = []
        total_parse = len(path_objects)
        decode_entities = self._should_decode_entities(base_path.suffix.lower())
        for idx, path in enumerate(path_objects, start=1):
            self._progress(
                f"Parsing version {idx}/{total_parse}...",
                0.1 + (idx / max(1, total_parse)) * 0.35,
            )
            try:
                document = parser.parse(str(path))
            except ParseError as exc:
                raise ParseError(str(path), exc.reason) from exc
            self._normalize_document_text_entities(
                document,
                decode_entities=decode_entities,
            )
            documents.append(document)

        comparisons: list[ComparisonResult] = []
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        total_compare = len(documents) - 1
        for idx in range(total_compare):
            step = idx + 1
            self._progress(
                f"Comparing version {step} to {step + 1}...",
                0.5 + (step / max(1, total_compare)) * 0.35,
            )
            comparisons.append(DiffEngine.compare(documents[idx], documents[idx + 1]))

        multi = MultiVersionResult(
            file_paths=files,
            comparisons=comparisons,
            documents=documents,
            report_paths=[],
        )

        self._progress("Generating summary report", 0.88)
        timestamp_label = datetime.now().strftime("%d-%m-%y--%H-%M-%S")
        summary_path = output_dir_path / f"versions_summary_{timestamp_label}.html"
        SummaryReporter().generate_versions(multi, str(summary_path))

        self._progress("Generating Excel report", 0.94)
        summary_excel_path = output_dir_path / f"versions_summary_{timestamp_label}.xlsx"
        multi.summary_excel_path = ExcelReporter().generate_versions(
            multi, str(summary_excel_path)
        )

        self._progress("Done", 1.0)
        return multi

    def compare_one_vs_all(
        self,
        reference_path: str,
        comparison_paths: list[str],
        output_dir: str,
    ) -> OneVsAllResult:
        ref_path = Path(reference_path)
        ref_ext = ref_path.suffix.lower()
        if ref_ext not in _ONE_VS_ALL_EXTENSIONS:
            raise UnsupportedFormatError(ref_ext)
        for path in comparison_paths:
            ext = Path(path).suffix.lower()
            if ext not in _ONE_VS_ALL_EXTENSIONS:
                raise UnsupportedFormatError(ext)

        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        self._progress("Selecting parser", 0.05)
        try:
            ref_parser = ParserRegistry.get_parser(str(ref_path))
        except UnsupportedFormatError as exc:
            raise UnsupportedFormatError(ref_ext) from exc

        decode_entities = self._should_decode_entities(ref_ext)

        self._progress("Parsing reference file", 0.1)
        try:
            ref_doc = ref_parser.parse(str(ref_path))
        except ParseError as exc:
            raise ParseError(reference_path, exc.reason) from exc
        self._normalize_document_text_entities(ref_doc, decode_entities=decode_entities)

        comparison_docs: list[ParsedDocument] = []
        total_files = len(comparison_paths)
        for idx, path in enumerate(comparison_paths, start=1):
            self._progress(
                f"Parsing comparison file {idx}/{total_files}...",
                0.1 + (idx / max(1, total_files + 1)) * 0.35,
            )
            try:
                cmp_parser = ParserRegistry.get_parser(path)
                cmp_doc = cmp_parser.parse(path)
            except ParseError as exc:
                raise ParseError(path, exc.reason) from exc
            self._normalize_document_text_entities(cmp_doc, decode_entities=decode_entities)
            comparison_docs.append(cmp_doc)

        comparisons: list[ComparisonResult] = []
        for idx, cmp_doc in enumerate(comparison_docs, start=1):
            self._progress(
                f"Comparing {idx}/{total_files}...",
                0.5 + (idx / max(1, total_files)) * 0.3,
            )
            comparisons.append(DiffEngine.compare(ref_doc, cmp_doc))

        result = OneVsAllResult(
            reference_path=reference_path,
            comparison_paths=comparison_paths,
            reference_doc=ref_doc,
            comparison_docs=comparison_docs,
            comparisons=comparisons,
        )

        self._progress("Generating HTML report", 0.83)
        timestamp_label = datetime.now().strftime("%d-%m-%y--%H-%M-%S")
        summary_path = output_dir_path / f"one_vs_all_{timestamp_label}.html"
        result.summary_html_path = SummaryReporter().generate_one_vs_all(
            result, str(summary_path)
        )

        self._progress("Generating Excel report", 0.94)
        summary_excel_path = output_dir_path / f"one_vs_all_{timestamp_label}.xlsx"
        result.summary_excel_path = ExcelReporter().generate_one_vs_all(
            result, str(summary_excel_path)
        )

        self._progress("Done", 1.0)
        return result

    def _progress(self, message: str, value: float) -> None:
        if self.on_progress is not None:
            self.on_progress(message, value)

    def _compare_pair_without_reports(
        self,
        file_a: str,
        file_b: str,
        *,
        excel_source_column_a: str | int | None = None,
        excel_source_column_b: str | int | None = None,
    ) -> ComparisonResult:
        path_a = Path(file_a)
        path_b = Path(file_b)
        ext_a = path_a.suffix.lower()
        ext_b = path_b.suffix.lower()
        if ext_a != ext_b:
            raise UnsupportedFormatError(f"{ext_a} vs {ext_b}")

        try:
            parser_a = ParserRegistry.get_parser(str(path_a))
            parser_b = ParserRegistry.get_parser(str(path_b))
        except UnsupportedFormatError as exc:
            raise UnsupportedFormatError(ext_a) from exc

        self._configure_excel_source_column(
            parser=parser_a,
            extension=ext_a,
            source_column=excel_source_column_a,
            file_path=file_a,
        )
        self._configure_excel_source_column(
            parser=parser_b,
            extension=ext_b,
            source_column=excel_source_column_b,
            file_path=file_b,
        )

        try:
            doc_a = parser_a.parse(str(path_a))
        except ParseError as exc:
            raise ParseError(file_a, exc.reason) from exc

        try:
            doc_b = parser_b.parse(str(path_b))
        except ParseError as exc:
            raise ParseError(file_b, exc.reason) from exc

        decode_entities = self._should_decode_entities(ext_a)
        self._normalize_document_text_entities(doc_a, decode_entities=decode_entities)
        self._normalize_document_text_entities(doc_b, decode_entities=decode_entities)

        return DiffEngine.compare(doc_a, doc_b)

    @staticmethod
    def _configure_excel_source_column(
        parser: BaseParser,
        extension: str,
        source_column: str | int | None,
        file_path: str,
    ) -> None:
        if extension not in {".xlsx", ".xls"}:
            return
        if source_column is None:
            return
        apply_source_column = getattr(parser, "set_source_column", None)
        if not callable(apply_source_column):
            return
        try:
            apply_source_column(source_column)
        except ValueError as exc:
            raise ParseError(file_path, str(exc)) from exc

    @staticmethod
    def _safe_stem(filename: str) -> str:
        safe = filename.replace(" ", "_")
        for bad in '<>:"/\\|?*':
            safe = safe.replace(bad, "_")
        return safe

    @staticmethod
    def _should_decode_entities(extension: str) -> bool:
        return extension.lower() not in {".txt", ".srt"}

    @staticmethod
    def _normalize_document_text_entities(
        document: ParsedDocument,
        *,
        decode_entities: bool,
    ) -> None:
        if not decode_entities:
            return

        from core.utils import decode_html_entities

        def normalize_value(value: Any) -> Any:
            if isinstance(value, str):
                return decode_html_entities(value)
            if isinstance(value, list):
                return [normalize_value(item) for item in value]
            if isinstance(value, tuple):
                return tuple(normalize_value(item) for item in value)
            if isinstance(value, dict):
                return {key: normalize_value(item) for key, item in value.items()}
            return value

        for segment in document.segments:
            segment.target = decode_html_entities(segment.target)
            if segment.source is not None:
                segment.source = decode_html_entities(segment.source)
            if segment.metadata:
                segment.metadata = normalize_value(segment.metadata)

        if document.metadata:
            document.metadata = normalize_value(document.metadata)
