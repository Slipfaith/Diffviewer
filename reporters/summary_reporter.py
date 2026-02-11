from __future__ import annotations

from difflib import SequenceMatcher
import html
from pathlib import Path
import re
from typing import Iterable

import xlsxwriter

from core.diff_engine import TextDiffer
from core.models import (
    BatchFileResult,
    BatchResult,
    ChunkType,
    MultiVersionResult,
    Segment,
)


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

            changes_headers = ["File", "Segment ID", "Source", "Old Target", "New Target"]
            changes_ws.write_row(0, 0, changes_headers, header)
            changes_ws.set_column(0, 0, 36)
            changes_ws.set_column(1, 1, 18)
            changes_ws.set_column(2, 4, 48)
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

        if not result.documents:
            return self._generate_versions_compact(result, output_file)

        return self._generate_versions_matrix(result, output_file)

    def _generate_versions_matrix(
        self,
        result: MultiVersionResult,
        output_file: Path,
    ) -> str:
        file_names = [Path(path).name for path in result.file_paths]
        rows = self._build_version_rows(result)
        changes_per_file = self._count_changes_per_file(rows, len(file_names))
        stat_blocks: list[str] = []
        for idx, name in enumerate(file_names):
            suffix = " (base)" if idx == 0 else ""
            stat_blocks.append(
                "<div class=\"stat\">"
                f"<div class=\"stat-label\">Changes in {self._escape(name)}{suffix}</div>"
                f"<div class=\"stat-value\">{changes_per_file[idx]}</div>"
                "</div>"
            )

        header_cells = [
            "<th class=\"col-segment-id\">Segment ID</th>",
            "<th>Source</th>",
        ]
        for name in file_names:
            header_cells.append(f"<th>Target: {self._escape(name)}</th>")

        body_rows: list[str] = []
        for row in rows:
            row_changed = self._row_has_changes(row)
            cells = [
                (
                    f"<td class=\"col-segment-id\" title=\"{self._escape(row['id'])}\">"
                    f"{self._escape(row['id'])}"
                    "</td>"
                ),
                f"<td>{self._escape_multiline(row['source'])}</td>",
            ]
            for idx, target in enumerate(row["targets"]):
                state = row["states"][idx]
                cell_classes = [f"state-{state}"]
                rendered_target = (
                    self._render_version_target(
                        previous_target=row["targets"][idx - 1],
                        current_target=target,
                        state=state,
                        version_index=idx,
                    )
                    if idx > 0
                    else self._escape_multiline(target)
                )
                cells.append(
                    f"<td class=\"{' '.join(cell_classes)}\">{rendered_target}</td>"
                )
            body_rows.append(
                f"<tr class=\"version-row\" data-changed=\"{'1' if row_changed else '0'}\">"
                + "".join(cells)
                + "</tr>"
            )

        extra_styles = """
.version-matrix { overflow-x: auto; }
.version-matrix table { min-width: 900px; }
.version-matrix .col-segment-id {
  width: 6ch;
  max-width: 6ch;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.version-matrix .state-base { color: #334155; }
.version-matrix .state-same { color: #111827; }
.version-matrix .state-missing { color: #94a3b8; font-style: italic; }
.version-matrix [class*="version-ins-"] {
  text-decoration: none;
  font-weight: 600;
}
.version-matrix [class*="version-del-"] {
  text-decoration: line-through;
  text-decoration-thickness: 1px;
}
.version-matrix .symbol-del {
  opacity: 0.82;
}
.version-matrix .ws-change {
  background: #e5e7eb;
  color: #111827;
  padding: 0 1px;
  border-radius: 3px;
}
"""
        version_color_styles = self._build_version_color_styles(len(file_names))
        script = """
(function () {
  const buttons = Array.from(document.querySelectorAll('.filter-btn[data-filter]'));
  const rows = Array.from(document.querySelectorAll('tr.version-row'));
  function applyFilter(mode) {
    rows.forEach((row) => {
      const changed = row.getAttribute('data-changed') === '1';
      row.classList.toggle('hidden', mode === 'changed' && !changed);
    });
    buttons.forEach((button) => {
      button.classList.toggle('active', button.getAttribute('data-filter') === mode);
    });
  }
  buttons.forEach((button) => {
    button.addEventListener('click', () => {
      const mode = button.getAttribute('data-filter') || 'changed';
      applyFilter(mode);
    });
  });
  applyFilter('changed');
})();
"""

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Version Matrix</title>
  <style>
{self._styles}
{extra_styles}
{version_color_styles}
  </style>
</head>
<body>
  <header class="page-header">
    <h1>Multi-Version Comparison (Stepwise)</h1>
  </header>
  <section class="stats">
    {''.join(stat_blocks)}
  </section>
  <section class="filters">
    <button class="filter-btn" type="button" data-filter="all">All text</button>
    <button class="filter-btn active" type="button" data-filter="changed">Changed</button>
  </section>
  <section class="version-matrix">
    <table class="report-table summary-table">
      <thead>
        <tr>{''.join(header_cells)}</tr>
      </thead>
      <tbody>
        {''.join(body_rows)}
      </tbody>
    </table>
  </section>
  <script>
{script}
  </script>
</body>
</html>"""

        output_file.write_text(html_content, encoding="utf-8")
        result.summary_report_path = str(output_file)
        return str(output_file)

    def _build_version_rows(self, result: MultiVersionResult) -> list[dict]:
        docs = result.documents
        doc_indices = [self._build_doc_index(doc.segments) for doc in docs]
        ordered_ids: list[str] = []
        seen_ids: set[str] = set()
        seen_source_keys: set[str] = set()

        if doc_indices:
            for segment in doc_indices[0]["segments"]:
                seen_ids.add(segment.id)
                ordered_ids.append(segment.id)
                source_key = self._compact_source(segment.source or "")
                if source_key:
                    seen_source_keys.add(source_key)

        for index in doc_indices[1:]:
            for segment in index["segments"]:
                if segment.id in seen_ids:
                    continue
                source_key = self._compact_source(segment.source or "")
                if source_key and source_key in seen_source_keys:
                    continue
                seen_ids.add(segment.id)
                ordered_ids.append(segment.id)
                if source_key:
                    seen_source_keys.add(source_key)

        rows: list[dict] = []
        for seg_id in ordered_ids:
            source = ""
            resolved_segments: list[Segment | None] = []

            for index in doc_indices:
                resolved_segments.append(index["by_id"].get(seg_id))

            source = self._first_non_empty_source(resolved_segments)
            targets: list[str] = []
            for index, segment in enumerate(resolved_segments):
                if segment is None and source:
                    segment = self._find_segment_by_source(doc_indices[index], source)
                if segment is None:
                    targets.append("")
                    continue
                target_value = getattr(segment, "target", "") or ""
                targets.append(target_value)
                if not source:
                    source_value = getattr(segment, "source", None)
                    if source_value:
                        source = source_value

            states = ["base"]
            for idx in range(1, len(targets)):
                previous_target = targets[idx - 1]
                target = targets[idx]
                if target == previous_target:
                    states.append("same")
                elif not previous_target and target:
                    states.append("added")
                elif previous_target and not target:
                    states.append("missing")
                else:
                    states.append("changed")

            rows.append(
                {
                    "id": seg_id,
                    "source": source,
                    "targets": targets,
                    "states": states,
                }
            )

        return rows

    @staticmethod
    def _build_doc_index(segments: list[Segment]) -> dict[str, object]:
        by_id: dict[str, Segment] = {}
        by_source: dict[str, list[Segment]] = {}
        by_compact_source: dict[str, list[Segment]] = {}

        for segment in segments:
            by_id.setdefault(segment.id, segment)
            source = (segment.source or "").strip()
            if not source:
                continue
            key = SummaryReporter._normalize_source(source)
            by_source.setdefault(key, []).append(segment)
            compact_key = SummaryReporter._compact_source(source)
            if compact_key:
                by_compact_source.setdefault(compact_key, []).append(segment)

        return {
            "segments": segments,
            "by_id": by_id,
            "by_source": by_source,
            "by_compact_source": by_compact_source,
        }

    def _find_segment_by_source(
        self,
        doc_index: dict[str, object],
        source: str,
    ) -> Segment | None:
        normalized = self._normalize_source(source)
        compact = self._compact_source(source)

        by_source = doc_index["by_source"]
        by_compact_source = doc_index["by_compact_source"]
        segments = doc_index["segments"]

        if normalized:
            exact = by_source.get(normalized, [])
            if exact:
                return self._pick_best_segment(exact)

        if compact:
            compact_matches = by_compact_source.get(compact, [])
            if compact_matches:
                return self._pick_best_segment(compact_matches)

        best: Segment | None = None
        best_score = 0.0
        for candidate in segments:
            candidate_source = candidate.source or ""
            if not candidate_source:
                continue
            score = self._source_similarity(source, candidate_source)
            if score > best_score:
                best_score = score
                best = candidate
        if best_score >= 0.72:
            return best
        return None

    @staticmethod
    def _pick_best_segment(candidates: list[Segment]) -> Segment:
        for candidate in candidates:
            if candidate.target:
                return candidate
        return candidates[0]

    @staticmethod
    def _first_non_empty_source(segments: list[Segment | None]) -> str:
        for segment in segments:
            if segment is not None and segment.source:
                return segment.source
        return ""

    @staticmethod
    def _normalize_source(value: str) -> str:
        return " ".join(value.casefold().split())

    @staticmethod
    def _compact_source(value: str) -> str:
        return "".join(re.findall(r"\w+", value.casefold(), flags=re.UNICODE))

    def _source_similarity(self, source_a: str, source_b: str) -> float:
        normalized_a = self._normalize_source(source_a)
        normalized_b = self._normalize_source(source_b)
        if not normalized_a or not normalized_b:
            return 0.0
        if normalized_a == normalized_b:
            return 1.0

        compact_a = self._compact_source(source_a)
        compact_b = self._compact_source(source_b)
        if compact_a and compact_b:
            if compact_a in compact_b or compact_b in compact_a:
                shorter = min(len(compact_a), len(compact_b))
                longer = max(len(compact_a), len(compact_b))
                if longer:
                    coverage = shorter / longer
                    return 0.82 + 0.18 * coverage

        return SequenceMatcher(None, normalized_a, normalized_b).ratio()

    def _build_version_color_styles(self, column_count: int) -> str:
        palette = [
            "#b45309",
            "#0369a1",
            "#6d28d9",
            "#15803d",
            "#be185d",
            "#4338ca",
        ]
        lines: list[str] = []
        for idx in range(1, max(1, column_count)):
            color = palette[(idx - 1) % len(palette)]
            lines.append(f".version-matrix .version-ins-{idx} {{ color: {color}; }}")
            lines.append(f".version-matrix .version-del-{idx} {{ color: {color}; }}")
        return "\n".join(lines)

    def _render_version_target(
        self,
        previous_target: str,
        current_target: str,
        state: str,
        version_index: int,
    ) -> str:
        if not current_target:
            return ""
        if state in {"base", "same"}:
            return self._escape_multiline(current_target)
        if state == "added":
            return self._wrap_version_insert(current_target, version_index)
        if TextDiffer.has_only_non_word_or_case_changes(previous_target, current_target):
            return self._render_non_text_diff(previous_target, current_target, version_index)

        diffs = TextDiffer.diff_auto(previous_target, current_target)
        parts: list[str] = []
        has_insertions = False
        for chunk in diffs:
            if chunk.type == ChunkType.EQUAL:
                parts.append(self._escape_multiline(chunk.text))
            elif chunk.type == ChunkType.INSERT:
                has_insertions = True
                parts.append(self._wrap_version_insert(chunk.text, version_index))
        if has_insertions:
            return "".join(parts)
        return self._escape_multiline(current_target)

    def _render_non_text_diff(
        self,
        previous_target: str,
        current_target: str,
        version_index: int,
    ) -> str:
        diffs = TextDiffer.diff_chars(previous_target, current_target)
        parts: list[str] = []
        for chunk in diffs:
            if chunk.type == ChunkType.EQUAL:
                parts.append(self._escape_multiline(chunk.text))
            elif chunk.type == ChunkType.INSERT:
                parts.append(
                    self._wrap_symbol_change(
                        chunk.text,
                        version_index=version_index,
                        is_deleted=False,
                    )
                )
            elif chunk.type == ChunkType.DELETE:
                parts.append(
                    self._wrap_symbol_change(
                        chunk.text,
                        version_index=version_index,
                        is_deleted=True,
                    )
                )
        return "".join(parts)

    def _wrap_version_insert(self, text: str, version_index: int) -> str:
        if not text:
            return ""
        whitespace_only = text.strip() == ""
        class_names = [f"version-ins-{version_index}"]
        if whitespace_only:
            class_names.append("ws-change")
        content = (
            self._escape_changed_text(text)
            if whitespace_only
            else self._escape_multiline(text)
        )
        return (
            f"<span class=\"{' '.join(class_names)}\">"
            f"{content}"
            "</span>"
        )

    def _wrap_symbol_change(
        self,
        text: str,
        version_index: int,
        is_deleted: bool,
    ) -> str:
        if not text:
            return ""
        whitespace_only = text.strip() == ""
        class_names = [
            f"version-del-{version_index}" if is_deleted else f"version-ins-{version_index}"
        ]
        if is_deleted:
            class_names.append("symbol-del")
        if whitespace_only:
            class_names.append("ws-change")
        content = (
            self._escape_changed_text(text)
            if whitespace_only
            else self._escape_multiline(text)
        )
        return f"<span class=\"{' '.join(class_names)}\">{content}</span>"

    @staticmethod
    def _row_has_changes(row: dict) -> bool:
        return any(state != "same" for state in row["states"][1:])

    @staticmethod
    def _count_changes_per_file(rows: list[dict], file_count: int) -> list[int]:
        counts = [0] * file_count
        for row in rows:
            states = row["states"]
            for idx in range(1, min(len(states), file_count)):
                if states[idx] != "same":
                    counts[idx] += 1
        return counts

    def _generate_versions_compact(
        self,
        result: MultiVersionResult,
        output_file: Path,
    ) -> str:
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

        html_content = f"""<!DOCTYPE html>
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

        output_file.write_text(html_content, encoding="utf-8")
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

    @staticmethod
    def _escape_multiline(value: str) -> str:
        return html.escape(value, quote=True).replace("\n", "<br>")

    @staticmethod
    def _escape_changed_text(value: str) -> str:
        return (
            html.escape(value, quote=True)
            .replace(" ", "&middot;")
            .replace("\n", "<br>")
        )

    def _stat_block(self, label: str, value: int) -> str:
        return (
            "<div class=\"stat\">"
            f"<div class=\"stat-label\">{self._escape(label)}</div>"
            f"<div class=\"stat-value\">{value}</div>"
            "</div>"
        )
