from __future__ import annotations

from pathlib import Path
from typing import Iterable

import xlsxwriter

from core.models import BatchFileResult, BatchResult, MultiVersionResult


class SummaryReporter:
    def __init__(self) -> None:
        styles_path = Path(__file__).resolve().parent / "templates" / "styles.css"
        base_styles = styles_path.read_text(encoding="utf-8") if styles_path.exists() else ""
        extra = """
.summary-table td a { color: #1d4ed8; text-decoration: none; }
.summary-table td a:hover { text-decoration: underline; }
.status-only { background: #fffbeb; }
.status-error { background: #fef2f2; }
.status-compared { background: #ffffff; }
"""
        self._styles = f"{base_styles}\n{extra}"

    def generate_batch(self, result: BatchResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != ".html":
            output_file = output_file.with_suffix(".html")
        output_file.parent.mkdir(parents=True, exist_ok=True)

        rows = [self._render_batch_row(item, output_file.parent) for item in result.files]

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Batch Comparison</title>
  <style>
{self._styles}
  </style>
</head>
<body>
  <header class="page-header">
    <h1>Batch Comparison: {self._escape(result.folder_a)} vs {self._escape(result.folder_b)}</h1>
  </header>
  <section class="stats">
    {self._stat_block("Total files", result.total_files)}
    {self._stat_block("Compared", result.compared_files)}
    {self._stat_block("Only in A", result.only_in_a)}
    {self._stat_block("Only in B", result.only_in_b)}
    {self._stat_block("Errors", result.errors)}
  </section>
  <table class="report-table summary-table">
    <thead>
      <tr>
        <th>File</th>
        <th>Status</th>
        <th>Added</th>
        <th>Deleted</th>
        <th>Modified</th>
        <th>Unchanged</th>
        <th>HTML report</th>
      </tr>
    </thead>
    <tbody>
      {''.join(rows)}
    </tbody>
  </table>
</body>
</html>"""

        output_file.write_text(html, encoding="utf-8")
        result.summary_report_path = str(output_file)
        return str(output_file)

    def generate_batch_excel(self, result: BatchResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != ".xlsx":
            output_file = output_file.with_suffix(".xlsx")
        output_file.parent.mkdir(parents=True, exist_ok=True)

        workbook = xlsxwriter.Workbook(
            str(output_file),
            {
                "strings_to_formulas": False,
                "strings_to_numbers": False,
                "strings_to_urls": False,
            },
        )
        try:
            summary_ws = workbook.add_worksheet("Summary")
            changes_ws = workbook.add_worksheet("All Changes")

            header = workbook.add_format({"bold": True, "bg_color": "#f3f4f6"})
            compared_row = workbook.add_format({"bg_color": "#ffffff", "text_wrap": True})
            only_row = workbook.add_format({"bg_color": "#fffbeb", "text_wrap": True})
            error_row = workbook.add_format({"bg_color": "#fef2f2", "text_wrap": True})
            wrap = workbook.add_format({"text_wrap": True})

            summary_headers = [
                "File",
                "Status",
                "Added",
                "Deleted",
                "Modified",
                "Unchanged",
                "HTML report",
            ]
            summary_ws.write_row(0, 0, summary_headers, header)
            summary_ws.set_column(0, 0, 36)
            summary_ws.set_column(1, 1, 14)
            summary_ws.set_column(2, 5, 12)
            summary_ws.set_column(6, 6, 48)
            summary_ws.freeze_panes(1, 0)
            summary_ws.autofilter(0, 0, max(1, len(result.files)), len(summary_headers) - 1)

            for row_index, item in enumerate(result.files, start=1):
                stats = item.statistics
                html_report = self._first_html_report(item.report_paths)
                row_format = compared_row
                if item.status in {"only_in_a", "only_in_b"}:
                    row_format = only_row
                elif item.status == "error":
                    row_format = error_row

                values = [
                    item.filename,
                    item.status,
                    stats.added if stats else "",
                    stats.deleted if stats else "",
                    stats.modified if stats else "",
                    stats.unchanged if stats else "",
                    html_report or item.error_message or "",
                ]
                for col_index, value in enumerate(values):
                    summary_ws.write_string(
                        row_index,
                        col_index,
                        "" if value is None else str(value),
                        row_format,
                    )

            changes_headers = ["File", "Segment ID", "Type", "Source", "Old Target", "New Target"]
            changes_ws.write_row(0, 0, changes_headers, header)
            changes_ws.set_column(0, 0, 36)
            changes_ws.set_column(1, 1, 18)
            changes_ws.set_column(2, 2, 12)
            changes_ws.set_column(3, 5, 48)
            changes_ws.freeze_panes(1, 0)

            row = 1
            for item in result.files:
                if item.comparison is None:
                    continue
                for change in item.comparison.changes:
                    before = change.segment_before
                    after = change.segment_after
                    segment_id = after.id if after is not None else before.id if before is not None else ""
                    source = after.source if after is not None else before.source if before is not None else ""
                    old_target = before.target if before is not None else ""
                    new_target = after.target if after is not None else ""
                    values = [
                        item.filename,
                        segment_id,
                        change.type.value,
                        source or "",
                        old_target,
                        new_target,
                    ]
                    for col_index, value in enumerate(values):
                        changes_ws.write_string(
                            row,
                            col_index,
                            "" if value is None else str(value),
                            wrap,
                        )
                    row += 1

            changes_ws.autofilter(0, 0, max(1, row - 1), len(changes_headers) - 1)
        finally:
            workbook.close()

        return str(output_file)

    def generate_versions(self, result: MultiVersionResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != ".html":
            output_file = output_file.with_suffix(".html")
        output_file.parent.mkdir(parents=True, exist_ok=True)

        rows = []
        for index, comparison in enumerate(result.comparisons):
            stats = comparison.statistics
            reports = result.report_paths[index] if index < len(result.report_paths) else []
            version_a = (
                Path(result.file_paths[index]).name
                if index < len(result.file_paths)
                else f"v{index + 1}"
            )
            version_b = (
                Path(result.file_paths[index + 1]).name
                if index + 1 < len(result.file_paths)
                else f"v{index + 2}"
            )
            rows.append(
                self._render_versions_row(
                    version_a=version_a,
                    version_b=version_b,
                    added=stats.added,
                    deleted=stats.deleted,
                    modified=stats.modified,
                    unchanged=stats.unchanged,
                    report_paths=reports,
                    base_dir=output_file.parent,
                )
            )

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Version History</title>
  <style>
{self._styles}
  </style>
</head>
<body>
  <header class="page-header">
    <h1>Version History</h1>
  </header>
  <table class="report-table summary-table">
    <thead>
      <tr>
        <th>Version A</th>
        <th>Version B</th>
        <th>Added</th>
        <th>Deleted</th>
        <th>Modified</th>
        <th>Unchanged</th>
        <th>HTML report</th>
      </tr>
    </thead>
    <tbody>
      {''.join(rows)}
    </tbody>
  </table>
</body>
</html>"""

        output_file.write_text(html, encoding="utf-8")
        result.summary_report_path = str(output_file)
        return str(output_file)

    def _render_batch_row(self, item: BatchFileResult, base_dir: Path) -> str:
        row_class = "status-compared"
        if item.status in {"only_in_a", "only_in_b"}:
            row_class = "status-only"
        elif item.status == "error":
            row_class = "status-error"

        stats = item.statistics
        added = str(stats.added) if stats else ""
        deleted = str(stats.deleted) if stats else ""
        modified = str(stats.modified) if stats else ""
        unchanged = str(stats.unchanged) if stats else ""

        html_report = self._first_html_report(item.report_paths)
        report_link = self._render_html_link(html_report, base_dir)
        if not report_link and item.error_message:
            report_link = self._escape(item.error_message)

        return (
            f"<tr class=\"{row_class}\">"
            f"<td>{self._escape(item.filename)}</td>"
            f"<td>{self._escape(item.status)}</td>"
            f"<td>{added}</td>"
            f"<td>{deleted}</td>"
            f"<td>{modified}</td>"
            f"<td>{unchanged}</td>"
            f"<td>{report_link}</td>"
            "</tr>"
        )

    def _render_versions_row(
        self,
        version_a: str,
        version_b: str,
        added: int,
        deleted: int,
        modified: int,
        unchanged: int,
        report_paths: list[str],
        base_dir: Path,
    ) -> str:
        html_report = self._first_html_report(report_paths)
        report_link = self._render_html_link(html_report, base_dir)
        return (
            "<tr>"
            f"<td>{self._escape(version_a)}</td>"
            f"<td>{self._escape(version_b)}</td>"
            f"<td>{added}</td>"
            f"<td>{deleted}</td>"
            f"<td>{modified}</td>"
            f"<td>{unchanged}</td>"
            f"<td>{report_link}</td>"
            "</tr>"
        )

    @staticmethod
    def _first_html_report(paths: Iterable[str]) -> str | None:
        for path in paths:
            if Path(path).suffix.lower() == ".html":
                return path
        return None

    def _render_html_link(self, path: str | None, base_dir: Path) -> str:
        if not path:
            return ""
        target = Path(path)
        try:
            rel = target.relative_to(base_dir)
        except ValueError:
            rel = Path(target.name)
        href = rel.as_posix()
        return f"<a href=\"{self._escape(href)}\">html</a>"

    @staticmethod
    def _escape(value: str) -> str:
        return (
            value.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def _stat_block(self, label: str, value: int) -> str:
        return (
            "<div class=\"stat\">"
            f"<div class=\"stat-label\">{self._escape(label)}</div>"
            f"<div class=\"stat-value\">{value}</div>"
            "</div>"
        )
