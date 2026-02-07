from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from core.diff_engine import DiffEngine
from core.models import ComparisonResult, ParseError, UnsupportedFormatError
from core.registry import ParserRegistry, ReporterRegistry
from reporters.excel_reporter import ExcelReporter
from reporters.html_reporter import HtmlReporter


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

        ext_label = ext_a.lstrip(".") or "unknown"
        base_name = f"report_{path_a.stem}_vs_{path_b.stem}_{ext_label}"

        reporters = [HtmlReporter(), ExcelReporter()]
        outputs: list[str] = []
        for reporter in reporters:
            output_path = output_dir_path / f"{base_name}{reporter.output_extension}"
            outputs.append(reporter.generate(result, str(output_path)))

        self._progress("Done", 1.0)
        return outputs

    def _progress(self, message: str, value: float) -> None:
        if self.on_progress is not None:
            self.on_progress(message, value)
