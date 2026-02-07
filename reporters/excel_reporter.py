from __future__ import annotations

from pathlib import Path

import xlsxwriter

from core.models import ChangeType, ChunkType, ComparisonResult, DiffChunk
from reporters.base import BaseReporter


class ExcelReporter(BaseReporter):
    name = "Excel Reporter"
    output_extension = ".xlsx"
    supports_rich_text = True

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        output_file.parent.mkdir(parents=True, exist_ok=True)
        workbook = xlsxwriter.Workbook(str(output_file))
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

            headers = ["#", "Segment ID", "Source", "Old Target", "New Target", "Type"]
            report_ws.write_row(0, 0, headers, header_format)

            report_ws.set_column(0, 0, 6)
            report_ws.set_column(1, 1, 15)
            report_ws.set_column(2, 2, 30)
            report_ws.set_column(3, 4, 45)
            report_ws.set_column(5, 5, 12)

            report_ws.freeze_panes(1, 0)
            end_row = max(0, len(result.changes))
            report_ws.autofilter(0, 0, end_row, 5)

            for index, change in enumerate(result.changes, start=1):
                row = index
                row_format = row_formats.get(change.type, row_formats[ChangeType.UNCHANGED])
                before = change.segment_before
                after = change.segment_after
                segment_id = after.id if after is not None else before.id if before is not None else ""
                source = after.source if after is not None else before.source if before is not None else ""

                report_ws.write(row, 0, index, row_format)
                report_ws.write(row, 1, segment_id, row_format)
                report_ws.write(row, 2, source or "", row_format)

                if change.type == ChangeType.MODIFIED:
                    self._write_rich(
                        report_ws,
                        row,
                        3,
                        change.text_diff,
                        row_format,
                        diff_formats,
                        side="old",
                    )
                    self._write_rich(
                        report_ws,
                        row,
                        4,
                        change.text_diff,
                        row_format,
                        diff_formats,
                        side="new",
                    )
                elif change.type == ChangeType.ADDED:
                    report_ws.write(row, 3, "", row_format)
                    report_ws.write(
                        row,
                        4,
                        after.target if after is not None else "",
                        row_format,
                    )
                elif change.type == ChangeType.DELETED:
                    report_ws.write(
                        row,
                        3,
                        before.target if before is not None else "",
                        row_format,
                    )
                    report_ws.write(row, 4, "", row_format)
                else:
                    report_ws.write(
                        row,
                        3,
                        before.target if before is not None else "",
                        row_format,
                    )
                    report_ws.write(
                        row,
                        4,
                        after.target if after is not None else "",
                        row_format,
                    )

                report_ws.write(row, 5, change.type.value.lower(), row_format)

            stats_labels = [
                ("Total", result.statistics.total_segments),
                ("Added", result.statistics.added),
                ("Deleted", result.statistics.deleted),
                ("Modified", result.statistics.modified),
                ("Unchanged", result.statistics.unchanged),
                ("Change %", f"{result.statistics.change_percentage * 100:.1f}%"),
            ]
            stats_ws.set_column(0, 0, 18)
            stats_ws.set_column(1, 1, 12)
            for row_index, (label, value) in enumerate(stats_labels):
                stats_ws.write(row_index, 0, label, header_format)
                stats_ws.write(row_index, 1, value)
        finally:
            workbook.close()

        return str(output_file)

    def _write_rich(
        self,
        worksheet,
        row: int,
        col: int,
        diffs: list[DiffChunk],
        cell_format,
        diff_formats: dict[ChunkType, object],
        side: str,
    ) -> None:
        fragments: list[object] = []
        text_buffer: list[str] = []

        for chunk in diffs:
            if chunk.type == ChunkType.EQUAL:
                text_buffer.append(chunk.text)
                continue
            if text_buffer:
                fragments.append("".join(text_buffer))
                text_buffer.clear()

            if chunk.type == ChunkType.DELETE and side == "old":
                fragments.append(diff_formats[ChunkType.DELETE])
                fragments.append(chunk.text)
            elif chunk.type == ChunkType.INSERT and side == "new":
                fragments.append(diff_formats[ChunkType.INSERT])
                fragments.append(chunk.text)

        if text_buffer:
            fragments.append("".join(text_buffer))

        plain_text = self._plain_text(diffs, side)
        string_fragments = [item for item in fragments if isinstance(item, str)]
        if len(string_fragments) < 2:
            worksheet.write(row, col, plain_text, cell_format)
            return

        try:
            worksheet.write_rich_string(row, col, *fragments, cell_format)
        except Exception:
            worksheet.write(row, col, plain_text, cell_format)

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
