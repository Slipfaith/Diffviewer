from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from jinja2 import Template

from core.models import ChangeType, ChunkType, ComparisonResult, DiffChunk
from reporters.base import BaseReporter


class HtmlReporter(BaseReporter):
    name = "HTML Reporter"
    output_extension = ".html"
    supports_rich_text = True

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        template_path = Path(__file__).resolve().parent / "templates" / "report.html.j2"
        styles_path = Path(__file__).resolve().parent / "templates" / "styles.css"
        template = Template(template_path.read_text(encoding="utf-8"))
        styles = styles_path.read_text(encoding="utf-8")

        rows = []
        for index, change in enumerate(result.changes, start=1):
            before = change.segment_before
            after = change.segment_after
            segment_id = after.id if after is not None else before.id if before is not None else ""
            source = after.source if after is not None else before.source if before is not None else ""
            is_changed = change.type != ChangeType.UNCHANGED
            rows.append(
                {
                    "index": index,
                    "segment_id": segment_id,
                    "source": source or "",
                    "old_target": self._render_old_target(change),
                    "new_target": self._render_new_target(change),
                    "change_type": change.type.value.lower(),
                    "is_changed": is_changed,
                }
            )

        html = template.render(
            styles=styles,
            file_a_name=Path(result.file_a.file_path).name,
            file_b_name=Path(result.file_b.file_path).name,
            timestamp=self._format_timestamp(result.timestamp),
            statistics=result.statistics,
            change_percentage=f"{result.statistics.change_percentage * 100:.1f}%",
            rows=rows,
        )

        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html, encoding="utf-8")
        return str(output_file)

    def _format_timestamp(self, timestamp: datetime) -> str:
        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=timezone.utc)
        return timestamp.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    def _render_old_target(self, change) -> str:
        if change.type == ChangeType.ADDED:
            return ""
        if change.type in (ChangeType.DELETED, ChangeType.UNCHANGED):
            return (change.segment_before.target if change.segment_before else "") or ""
        return self._render_diff(change.text_diff, side="old")

    def _render_new_target(self, change) -> str:
        if change.type == ChangeType.DELETED:
            return ""
        if change.type in (ChangeType.ADDED, ChangeType.UNCHANGED):
            return (change.segment_after.target if change.segment_after else "") or ""
        return self._render_diff(change.text_diff, side="new")

    @staticmethod
    def _escape(text: str) -> str:
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def _render_diff(self, diffs: list[DiffChunk], side: str) -> str:
        parts: list[str] = []
        for chunk in diffs:
            text = self._escape(chunk.text)
            if chunk.type == ChunkType.EQUAL:
                parts.append(text)
            elif chunk.type == ChunkType.DELETE and side == "old":
                parts.append(f"<del>{text}</del>")
            elif chunk.type == ChunkType.INSERT and side == "new":
                parts.append(f"<ins>{text}</ins>")
        return "".join(parts)
