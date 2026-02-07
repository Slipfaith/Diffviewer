from __future__ import annotations

from pathlib import Path
from typing import Iterable

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
.timeline { margin: 12px 0 20px; padding: 12px; background: #ffffff; border-radius: 10px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
.timeline-item { padding: 6px 0; }
"""
        self._styles = f"{base_styles}\n{extra}"

    def generate_batch(self, result: BatchResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != ".html":
            output_file = output_file.with_suffix(".html")
        output_file.parent.mkdir(parents=True, exist_ok=True)

        rows = []
        for item in result.files:
            rows.append(self._render_batch_row(item, output_file.parent))

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
        <th>Reports</th>
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

    def generate_versions(self, result: MultiVersionResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != ".html":
            output_file = output_file.with_suffix(".html")
        output_file.parent.mkdir(parents=True, exist_ok=True)

        base_name = Path(result.file_paths[0]).name if result.file_paths else "versions"
        timeline_items = []
        for index, comparison in enumerate(result.comparisons):
            stats = comparison.statistics
            timeline_items.append(
                f"<div class=\"timeline-item\">v{index + 1} → v{index + 2}: "
                f"added {stats.added}, deleted {stats.deleted}, modified {stats.modified}</div>"
            )

        rows = []
        for index, comparison in enumerate(result.comparisons):
            stats = comparison.statistics
            reports = result.report_paths[index] if index < len(result.report_paths) else []
            rows.append(
                self._render_versions_row(
                    f"v{index + 1}",
                    f"v{index + 2}",
                    stats.added,
                    stats.deleted,
                    stats.modified,
                    reports,
                    output_file.parent,
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
    <h1>Version History: {self._escape(base_name)}</h1>
  </header>
  <section class="timeline">
    {''.join(timeline_items)}
  </section>
  <table class="report-table summary-table">
    <thead>
      <tr>
        <th>Version A</th>
        <th>Version B</th>
        <th>Added</th>
        <th>Deleted</th>
        <th>Modified</th>
        <th>Reports</th>
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
        status = item.status
        row_class = "status-compared"
        if status in {"only_in_a", "only_in_b"}:
            row_class = "status-only"
        elif status == "error":
            row_class = "status-error"

        stats = item.statistics
        added = stats.added if stats else ""
        deleted = stats.deleted if stats else ""
        modified = stats.modified if stats else ""
        unchanged = stats.unchanged if stats else ""

        report_links = self._render_links(item.report_paths, base_dir)
        return (
            f"<tr class=\"{row_class}\">"
            f"<td>{self._escape(item.filename)}</td>"
            f"<td>{self._escape(status)}</td>"
            f"<td>{added}</td>"
            f"<td>{deleted}</td>"
            f"<td>{modified}</td>"
            f"<td>{unchanged}</td>"
            f"<td>{report_links or self._escape(item.error_message or '')}</td>"
            "</tr>"
        )

    def _render_versions_row(
        self,
        version_a: str,
        version_b: str,
        added: int,
        deleted: int,
        modified: int,
        report_paths: list[str],
        base_dir: Path,
    ) -> str:
        report_links = self._render_links(report_paths, base_dir)
        return (
            "<tr>"
            f"<td>{self._escape(version_a)}</td>"
            f"<td>{self._escape(version_b)}</td>"
            f"<td>{added}</td>"
            f"<td>{deleted}</td>"
            f"<td>{modified}</td>"
            f"<td>{report_links}</td>"
            "</tr>"
        )

    @staticmethod
    def _escape(value: str) -> str:
        return (
            value.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def _render_links(self, paths: Iterable[str], base_dir: Path) -> str:
        links = []
        for path in paths:
            if not path:
                continue
            rel = Path(path)
            try:
                rel_path = rel.relative_to(base_dir)
            except ValueError:
                rel_path = Path(Path(path).name)
            label = rel_path.suffix.lstrip(".") or "file"
            links.append(f"<a href=\"{rel_path.as_posix()}\">{label}</a>")
        return " | ".join(links)

    def _stat_block(self, label: str, value: int) -> str:
        return (
            "<div class=\"stat\">"
            f"<div class=\"stat-label\">{self._escape(label)}</div>"
            f"<div class=\"stat-value\">{value}</div>"
            "</div>"
        )
