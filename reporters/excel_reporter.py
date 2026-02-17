from __future__ import annotations

import json
import re
from pathlib import Path

import xlsxwriter

from core.models import ChangeStatistics, ChangeType, ChunkType, ComparisonResult, DiffChunk
from core.utils import decode_html_entities
from reporters.base import BaseReporter


class ExcelReporter(BaseReporter):
    name = "Excel Reporter"
    output_extension = ".xlsx"
    supports_rich_text = True

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        payload = {
            "file_a_name": Path(result.file_a.file_path).name,
            "file_b_name": Path(result.file_b.file_path).name,
            "statistics": {
                "total_segments": result.statistics.total_segments,
                "added": result.statistics.added,
                "deleted": result.statistics.deleted,
                "modified": result.statistics.modified,
                "unchanged": result.statistics.unchanged,
                "change_percentage": result.statistics.change_percentage,
            },
            "changes": [
                {
                    "type": change.type.value,
                    "segment_before": self._serialize_segment(change.segment_before),
                    "segment_after": self._serialize_segment(change.segment_after),
                    "text_diff": [
                        {"type": chunk.type.value, "text": chunk.text}
                        for chunk in change.text_diff
                    ],
                }
                for change in result.changes
            ],
        }
        return self.generate_from_json(payload, output_path)

    def generate_multi(
        self,
        comparisons: list[tuple[str, ComparisonResult]],
        output_path: str,
    ) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

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
            report_ws = workbook.add_worksheet("All Changes")
            stats_ws = workbook.add_worksheet("Statistics")

            header_format = workbook.add_format(
                {"bold": True, "bg_color": "#f3f4f6", "text_wrap": True}
            )
            row_formats = {
                ChangeType.ADDED: workbook.add_format(
                    {"bg_color": "#ecfdf3", "text_wrap": True}
                ),
                ChangeType.DELETED: workbook.add_format(
                    {"bg_color": "#fef2f2", "text_wrap": True}
                ),
                ChangeType.MODIFIED: workbook.add_format(
                    {"bg_color": "#fffbeb", "text_wrap": True}
                ),
                ChangeType.UNCHANGED: workbook.add_format({"text_wrap": True}),
                ChangeType.MOVED: workbook.add_format(
                    {"bg_color": "#eef2ff", "text_wrap": True}
                ),
            }

            show_source = any(
                bool(
                    (
                        (change.segment_after.source if change.segment_after is not None else None)
                        or (change.segment_before.source if change.segment_before is not None else None)
                        or ""
                    ).strip()
                )
                for _, comparison in comparisons
                for change in comparison.changes
            )

            headers = ["#", "File Pair", "File A", "File B", "Segment ID"]
            if show_source:
                headers.append("Source")
            headers.extend(["Old Target", "New Target"])
            report_ws.write_row(0, 0, headers, header_format)

            col_index = 0
            col_pair = 1
            col_file_a = 2
            col_file_b = 3
            col_segment = 4
            if show_source:
                col_source: int | None = 5
                col_old = 6
                col_new = 7
                report_ws.set_column(0, 0, 6)
                report_ws.set_column(1, 1, 34)
                report_ws.set_column(2, 3, 24)
                report_ws.set_column(4, 4, 18)
                report_ws.set_column(5, 5, 30)
                report_ws.set_column(6, 7, 46)
            else:
                col_source = None
                col_old = 5
                col_new = 6
                report_ws.set_column(0, 0, 6)
                report_ws.set_column(1, 1, 34)
                report_ws.set_column(2, 3, 24)
                report_ws.set_column(4, 4, 18)
                report_ws.set_column(5, 6, 46)
            report_ws.freeze_panes(1, 0)

            all_changes = []
            row = 1
            sequence = 1
            for pair_label, comparison in comparisons:
                file_a_name = Path(comparison.file_a.file_path).name
                file_b_name = Path(comparison.file_b.file_path).name
                for change in comparison.changes:
                    all_changes.append(change)
                    row_format = row_formats.get(change.type, row_formats[ChangeType.UNCHANGED])
                    before = change.segment_before
                    after = change.segment_after
                    segment_id = after.id if after is not None else before.id if before is not None else ""
                    source = (
                        (after.source if after is not None else None)
                        or (before.source if before is not None else None)
                        or ""
                    )
                    old_target = before.target if before is not None else ""
                    new_target = after.target if after is not None else ""

                    report_ws.write_number(row, col_index, sequence, row_format)
                    self._write_text(report_ws, row, col_pair, pair_label, row_format)
                    self._write_text(report_ws, row, col_file_a, file_a_name, row_format)
                    self._write_text(report_ws, row, col_file_b, file_b_name, row_format)
                    self._write_text(report_ws, row, col_segment, segment_id, row_format)
                    if col_source is not None:
                        self._write_text(report_ws, row, col_source, source, row_format)
                    self._write_text(report_ws, row, col_old, old_target, row_format)
                    self._write_text(report_ws, row, col_new, new_target, row_format)
                    if change.type == ChangeType.UNCHANGED:
                        report_ws.set_row(row, None, None, {"hidden": True})
                    row += 1
                    sequence += 1

            end_row = max(1, row - 1)
            report_ws.autofilter(0, 0, end_row, col_new)

            statistics = ChangeStatistics.from_changes(all_changes)
            stats_labels = [
                ("Pairs", len(comparisons)),
                ("Total", statistics.total_segments),
                ("Added", statistics.added),
                ("Deleted", statistics.deleted),
                ("Modified", statistics.modified),
                ("Moved", statistics.moved),
                ("Unchanged", statistics.unchanged),
                ("Change %", f"{statistics.change_percentage * 100:.1f}%"),
            ]
            stats_ws.set_column(0, 0, 18)
            stats_ws.set_column(1, 1, 14)
            for row_index, (label, value) in enumerate(stats_labels):
                stats_ws.write(row_index, 0, label, header_format)
                stats_ws.write(row_index, 1, value)
        finally:
            workbook.close()

        return str(output_file)

    def generate_from_html(self, html_path: str, output_path: str | None = None) -> str:
        html_file = Path(html_path)
        content = html_file.read_text(encoding="utf-8")
        data = self._extract_report_data(content)
        if output_path is None:
            output_file = html_file.with_suffix(self.output_extension)
        else:
            output_file = Path(output_path)
        return self.generate_from_json(data, str(output_file))

    def generate_from_json(self, data: dict, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

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
            report_ws = workbook.add_worksheet("Report")
            stats_ws = workbook.add_worksheet("Statistics")

            header_format = workbook.add_format(
                {"bold": True, "bg_color": "#f3f4f6", "text_wrap": True}
            )
            row_formats = {
                ChangeType.ADDED: workbook.add_format(
                    {"bg_color": "#ecfdf3", "text_wrap": True}
                ),
                ChangeType.DELETED: workbook.add_format(
                    {"bg_color": "#fef2f2", "text_wrap": True}
                ),
                ChangeType.MODIFIED: workbook.add_format(
                    {"bg_color": "#fffbeb", "text_wrap": True}
                ),
                ChangeType.UNCHANGED: workbook.add_format({"text_wrap": True}),
                ChangeType.MOVED: workbook.add_format(
                    {"bg_color": "#eef2ff", "text_wrap": True}
                ),
            }
            diff_formats = {
                ChunkType.DELETE: workbook.add_format(
                    {"font_color": "#b91c1c", "font_strikeout": True, "text_wrap": True}
                ),
                ChunkType.INSERT: workbook.add_format(
                    {"font_color": "#15803d", "underline": True, "text_wrap": True}
                ),
                ChunkType.EQUAL: workbook.add_format({"text_wrap": True}),
            }
            delete_cell_formats = {
                ChangeType.ADDED: workbook.add_format(
                    {
                        "bg_color": "#ecfdf3",
                        "text_wrap": True,
                        "font_color": "#b91c1c",
                        "font_strikeout": True,
                    }
                ),
                ChangeType.DELETED: workbook.add_format(
                    {
                        "bg_color": "#fef2f2",
                        "text_wrap": True,
                        "font_color": "#b91c1c",
                        "font_strikeout": True,
                    }
                ),
                ChangeType.MODIFIED: workbook.add_format(
                    {
                        "bg_color": "#fffbeb",
                        "text_wrap": True,
                        "font_color": "#b91c1c",
                        "font_strikeout": True,
                    }
                ),
                ChangeType.UNCHANGED: workbook.add_format(
                    {"text_wrap": True, "font_color": "#b91c1c", "font_strikeout": True}
                ),
                ChangeType.MOVED: workbook.add_format(
                    {
                        "bg_color": "#eef2ff",
                        "text_wrap": True,
                        "font_color": "#b91c1c",
                        "font_strikeout": True,
                    }
                ),
            }
            insert_cell_formats = {
                ChangeType.ADDED: workbook.add_format(
                    {
                        "bg_color": "#ecfdf3",
                        "text_wrap": True,
                        "font_color": "#15803d",
                        "underline": True,
                    }
                ),
                ChangeType.DELETED: workbook.add_format(
                    {
                        "bg_color": "#fef2f2",
                        "text_wrap": True,
                        "font_color": "#15803d",
                        "underline": True,
                    }
                ),
                ChangeType.MODIFIED: workbook.add_format(
                    {
                        "bg_color": "#fffbeb",
                        "text_wrap": True,
                        "font_color": "#15803d",
                        "underline": True,
                    }
                ),
                ChangeType.UNCHANGED: workbook.add_format(
                    {"text_wrap": True, "font_color": "#15803d", "underline": True}
                ),
                ChangeType.MOVED: workbook.add_format(
                    {
                        "bg_color": "#eef2ff",
                        "text_wrap": True,
                        "font_color": "#15803d",
                        "underline": True,
                    }
                ),
            }

            show_source = self._should_show_source(data)
            headers = ["#", "Segment ID"]
            if show_source:
                headers.append("Source")
            headers.extend(["Old Target", "New Target"])
            report_ws.write_row(0, 0, headers, header_format)
            col_index = 0
            col_segment = 1
            if show_source:
                col_source: int | None = 2
                col_old = 3
                col_new = 4
                report_ws.set_column(0, 0, 6)
                report_ws.set_column(1, 1, 15)
                report_ws.set_column(2, 2, 30)
                report_ws.set_column(3, 4, 45)
            else:
                col_source = None
                col_old = 2
                col_new = 3
                report_ws.set_column(0, 0, 6)
                report_ws.set_column(1, 1, 15)
                report_ws.set_column(2, 3, 45)
            report_ws.freeze_panes(1, 0)
            end_row = max(0, len(data.get("changes", [])))
            report_ws.autofilter(0, 0, end_row, col_new)

            for index, change in enumerate(data.get("changes", []), start=1):
                row = index
                change_type = self._parse_change_type(
                    change.get("type", ChangeType.UNCHANGED.value)
                )
                row_format = row_formats.get(change_type, row_formats[ChangeType.UNCHANGED])

                before = change.get("segment_before") or {}
                after = change.get("segment_after") or {}
                segment_id = after.get("id") or before.get("id") or ""
                source = after.get("source") or before.get("source") or ""
                text_diff = [
                    DiffChunk(
                        type=self._parse_chunk_type(chunk.get("type", ChunkType.EQUAL.value)),
                        text=chunk.get("text", ""),
                    )
                    for chunk in change.get("text_diff", [])
                ]

                report_ws.write_number(row, col_index, index, row_format)
                self._write_text(report_ws, row, col_segment, segment_id, row_format)
                if col_source is not None:
                    self._write_text(report_ws, row, col_source, source, row_format)

                if change_type == ChangeType.MODIFIED:
                    old_target = before.get("target") if before else ""
                    new_target = after.get("target") if after else ""
                    self._write_rich(
                        report_ws,
                        row,
                        col_old,
                        text_diff,
                        row_format,
                        diff_formats,
                        side="old",
                        fallback=old_target,
                    )
                    self._write_rich(
                        report_ws,
                        row,
                        col_new,
                        text_diff,
                        row_format,
                        diff_formats,
                        side="new",
                        fallback=new_target,
                    )
                elif change_type == ChangeType.ADDED:
                    self._write_text(report_ws, row, col_old, "", row_format)
                    self._write_text(
                        report_ws,
                        row,
                        col_new,
                        after.get("target", ""),
                        insert_cell_formats[change_type],
                    )
                elif change_type == ChangeType.DELETED:
                    self._write_text(
                        report_ws,
                        row,
                        col_old,
                        before.get("target", ""),
                        delete_cell_formats[change_type],
                    )
                    self._write_text(report_ws, row, col_new, "", row_format)
                else:
                    self._write_text(
                        report_ws,
                        row,
                        col_old,
                        before.get("target", ""),
                        row_format,
                    )
                    self._write_text(
                        report_ws,
                        row,
                        col_new,
                        after.get("target", ""),
                        row_format,
                    )

                if change_type == ChangeType.UNCHANGED:
                    report_ws.set_row(row, None, None, {"hidden": True})

            stats = data.get("statistics", {})
            stats_labels = [
                ("Total", stats.get("total_segments", len(data.get("changes", [])))),
                ("Added", stats.get("added", 0)),
                ("Deleted", stats.get("deleted", 0)),
                ("Modified", stats.get("modified", 0)),
                ("Unchanged", stats.get("unchanged", 0)),
                (
                    "Change %",
                    f"{float(stats.get('change_percentage', 0.0)) * 100:.1f}%",
                ),
            ]
            stats_ws.set_column(0, 0, 18)
            stats_ws.set_column(1, 1, 12)
            for row_index, (label, value) in enumerate(stats_labels):
                stats_ws.write(row_index, 0, label, header_format)
                stats_ws.write(row_index, 1, value)
        finally:
            workbook.close()

        return str(output_file)

    @staticmethod
    def _serialize_segment(segment) -> dict | None:
        if segment is None:
            return None
        return {
            "id": segment.id,
            "source": segment.source,
            "target": segment.target,
            "metadata": segment.metadata,
        }

    @staticmethod
    def _extract_report_data(html_content: str) -> dict:
        match = re.search(
            r'<script[^>]*id="report-data"[^>]*type="application/json"[^>]*>(.*?)</script>',
            html_content,
            flags=re.DOTALL | re.IGNORECASE,
        )
        if not match:
            raise ValueError("report-data JSON block not found in HTML")
        return json.loads(match.group(1).strip())

    @staticmethod
    def _parse_change_type(raw_value: str) -> ChangeType:
        value = (raw_value or "").upper()
        try:
            return ChangeType[value]
        except KeyError as exc:
            raise ValueError(f"Unsupported change type: {raw_value}") from exc

    @staticmethod
    def _parse_chunk_type(raw_value: str) -> ChunkType:
        value = (raw_value or "").upper()
        try:
            return ChunkType[value]
        except KeyError as exc:
            raise ValueError(f"Unsupported chunk type: {raw_value}") from exc

    @staticmethod
    def _should_show_source(data: dict) -> bool:
        for change in data.get("changes", []):
            before = change.get("segment_before") or {}
            after = change.get("segment_after") or {}
            source = after.get("source") or before.get("source") or ""
            if str(source).strip():
                return True
        return False

    def _write_rich(
        self,
        worksheet,
        row: int,
        col: int,
        diffs: list[DiffChunk],
        cell_format,
        diff_formats: dict[ChunkType, object],
        side: str,
        fallback: str = "",
    ) -> None:
        fragments: list[object] = []
        text_buffer: list[str] = []

        for chunk in diffs:
            chunk_text = decode_html_entities(
                chunk.text,
                decode_single_encoded=False,
            )
            if chunk.type == ChunkType.EQUAL:
                text_buffer.append(chunk_text)
                continue
            if text_buffer:
                fragments.append("".join(text_buffer))
                text_buffer.clear()

            if chunk.type == ChunkType.DELETE and side == "old":
                fragments.append(diff_formats[ChunkType.DELETE])
                fragments.append(chunk_text)
            elif chunk.type == ChunkType.INSERT and side == "new":
                fragments.append(diff_formats[ChunkType.INSERT])
                fragments.append(chunk_text)

        if text_buffer:
            fragments.append("".join(text_buffer))

        plain_text = self._plain_text(diffs, side)
        if not plain_text:
            plain_text = fallback if fallback is not None else ""
        plain_text = str(plain_text)

        # XlsxWriter requires more than two rich fragments and emits a warning
        # otherwise. Also skip rich-string write when we have no styled runs.
        has_rich_runs = any(not isinstance(item, str) for item in fragments)
        if len(fragments) <= 2 or not has_rich_runs:
            self._write_text(worksheet, row, col, plain_text, cell_format)
            return

        try:
            result = worksheet.write_rich_string(row, col, *fragments, cell_format)
            if result != 0:
                self._write_text(worksheet, row, col, plain_text, cell_format)
        except Exception:
            self._write_text(worksheet, row, col, plain_text, cell_format)

    @staticmethod
    def _write_text(worksheet, row: int, col: int, value: object, cell_format) -> None:
        text = "" if value is None else str(value)
        text = decode_html_entities(text, decode_single_encoded=False)
        worksheet.write_string(row, col, text, cell_format)

    @staticmethod
    def _plain_text(diffs: list[DiffChunk], side: str) -> str:
        parts: list[str] = []
        for chunk in diffs:
            if chunk.type == ChunkType.EQUAL:
                parts.append(chunk.text)
            elif chunk.type == ChunkType.DELETE and side == "old":
                parts.append(chunk.text)
            elif chunk.type == ChunkType.INSERT and side == "new":
                parts.append(chunk.text)
        return "".join(parts)
