from __future__ import annotations

from datetime import datetime, timezone
import html
from pathlib import Path

from jinja2 import Template

from core.models import ChangeStatistics, ChangeType, ChunkType, ComparisonResult, DiffChunk
from core.utils import resource_path
from reporters.base import BaseReporter


class HtmlReporter(BaseReporter):
    name = "HTML Reporter"
    output_extension = ".html"
    supports_rich_text = True

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = self._normalize_output_path(output_path)
        template, styles = self._load_template_assets()
        show_source = any(bool(self._source_text(change).strip()) for change in result.changes)
        rows = self._build_rows(
            result.changes,
            start_index=1,
            file_key="file-1",
            file_label=f"{Path(result.file_a.file_path).name} vs {Path(result.file_b.file_path).name}",
        )
        html_content = template.render(
            styles=styles,
            report_title="Change Tracker",
            report_subtitle=f"{Path(result.file_a.file_path).name} vs {Path(result.file_b.file_path).name}",
            file_a_name=Path(result.file_a.file_path).name,
            file_b_name=Path(result.file_b.file_path).name,
            timestamp=self._format_timestamp(result.timestamp),
            statistics=result.statistics,
            change_percentage=f"{result.statistics.change_percentage * 100:.1f}%",
            rows=rows,
            show_source=show_source,
            multi_mode=False,
            file_options=[],
        )

        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding="utf-8")
        return str(output_file)

    def generate_multi(
        self,
        comparisons: list[tuple[str, ComparisonResult]],
        output_path: str,
    ) -> str:
        output_file = self._normalize_output_path(output_path)
        template, styles = self._load_template_assets()

        rows: list[dict[str, object]] = []
        file_options: list[dict[str, str]] = []
        all_changes = []
        show_source = False
        next_row_index = 1
        timestamps = []

        for index, (file_label, comparison) in enumerate(comparisons, start=1):
            file_key = f"file-{index}"
            file_options.append(
                {
                    "key": file_key,
                    "label": self._escape(file_label),
                }
            )
            rows.extend(
                self._build_rows(
                    comparison.changes,
                    start_index=next_row_index,
                    file_key=file_key,
                    file_label=file_label,
                )
            )
            next_row_index += len(comparison.changes)
            all_changes.extend(comparison.changes)
            if not show_source:
                show_source = any(
                    bool(self._source_text(change).strip()) for change in comparison.changes
                )
            timestamps.append(self._ensure_utc_timestamp(comparison.timestamp))

        statistics = ChangeStatistics.from_changes(all_changes)
        report_timestamp = max(timestamps) if timestamps else datetime.now(timezone.utc)
        html_content = template.render(
            styles=styles,
            report_title="Change Tracker",
            report_subtitle=f"Combined report for {len(comparisons)} file pairs",
            file_a_name="",
            file_b_name="",
            timestamp=self._format_timestamp(report_timestamp),
            statistics=statistics,
            change_percentage=f"{statistics.change_percentage * 100:.1f}%",
            rows=rows,
            show_source=show_source,
            multi_mode=True,
            file_options=file_options,
        )

        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding="utf-8")
        return str(output_file)

    def _format_timestamp(self, timestamp: datetime) -> str:
        timestamp_utc = self._ensure_utc_timestamp(timestamp)
        return timestamp_utc.strftime("%Y-%m-%d %H:%M:%S UTC")

    def _ensure_utc_timestamp(self, timestamp: datetime) -> datetime:
        if timestamp.tzinfo is None:
            return timestamp.replace(tzinfo=timezone.utc)
        return timestamp.astimezone(timezone.utc)

    def _normalize_output_path(self, output_path: str) -> Path:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)
        return output_file

    def _load_template_assets(self) -> tuple[Template, str]:
        template_dir = Path(resource_path("reporters/templates"))
        template_path = template_dir / "report.html.j2"
        styles_path = template_dir / "styles.css"
        template = Template(template_path.read_text(encoding="utf-8"))
        styles = styles_path.read_text(encoding="utf-8")
        return template, styles

    def _build_rows(
        self,
        changes,
        *,
        start_index: int,
        file_key: str,
        file_label: str,
    ) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for offset, change in enumerate(changes):
            before = change.segment_before
            after = change.segment_after
            segment_id = after.id if after is not None else before.id if before is not None else ""
            source = self._source_text(change)
            is_changed = change.type != ChangeType.UNCHANGED
            rows.append(
                {
                    "index": start_index + offset,
                    "file_key": file_key,
                    "file_label": self._escape(file_label),
                    "segment_id": segment_id,
                    "source": self._escape(source or ""),
                    "old_target": self._render_old_target(change),
                    "new_target": self._render_new_target(change),
                    "row_state": "changed" if is_changed else "unchanged",
                    "is_changed": is_changed,
                }
            )
        return rows

    def _render_old_target(self, change) -> str:
        if change.type == ChangeType.ADDED:
            return ""
        if change.type == ChangeType.DELETED:
            text = (change.segment_before.target if change.segment_before else "") or ""
            return self._wrap_delete(text)
        if change.type == ChangeType.UNCHANGED:
            return self._escape((change.segment_before.target if change.segment_before else "") or "")
        return self._render_diff(change.text_diff, side="old")

    def _render_new_target(self, change) -> str:
        if change.type == ChangeType.DELETED:
            return ""
        if change.type == ChangeType.ADDED:
            text = (change.segment_after.target if change.segment_after else "") or ""
            return self._wrap_insert(text)
        if change.type == ChangeType.UNCHANGED:
            return self._escape((change.segment_after.target if change.segment_after else "") or "")
        rendered = self._render_diff(change.text_diff, side="new")
        if rendered:
            return rendered
        fallback = (change.segment_after.target if change.segment_after else "") or ""
        return self._escape(fallback)

    @staticmethod
    def _source_text(change) -> str:
        before = change.segment_before
        after = change.segment_after
        source = after.source if after is not None else before.source if before is not None else ""
        return source or ""

    @staticmethod
    def _escape(text: str) -> str:
        return html.escape(text, quote=True)

    def _escape_changed_text(self, text: str) -> str:
        return self._escape(text).replace(" ", "&middot;")

    def _wrap_delete(self, text: str) -> str:
        return f"<del>{self._escape_changed_text(text)}</del>" if text else ""

    def _wrap_insert(self, text: str) -> str:
        return f"<ins>{self._escape_changed_text(text)}</ins>" if text else ""

    def _render_diff(self, diffs: list[DiffChunk], side: str) -> str:
        parts: list[str] = []
        for chunk in diffs:
            text = self._escape(chunk.text)
            if chunk.type == ChunkType.EQUAL:
                parts.append(text)
            elif chunk.type == ChunkType.DELETE and side == "old":
                parts.append(self._wrap_delete(chunk.text))
            elif chunk.type == ChunkType.INSERT and side == "new":
                parts.append(self._wrap_insert(chunk.text))
        return "".join(parts)
