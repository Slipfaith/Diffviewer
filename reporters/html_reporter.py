from __future__ import annotations

from datetime import datetime, timezone
import html
from pathlib import Path

from jinja2 import Template

from core.models import ChangeType, ChunkType, ComparisonResult, DiffChunk
from core.utils import resource_path
from reporters.base import BaseReporter


class HtmlReporter(BaseReporter):
    name = "HTML Reporter"
    output_extension = ".html"
    supports_rich_text = True

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        template_dir = Path(resource_path("reporters/templates"))
        template_path = template_dir / "report.html.j2"
        styles_path = template_dir / "styles.css"
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
                    "source": self._escape(source or ""),
                    "old_target": self._render_old_target(change),
                    "new_target": self._render_new_target(change),
                    "change_type": change.type.value.lower(),
                    "is_changed": is_changed,
                }
            )

        html_content = template.render(
            styles=styles,
            file_a_name=Path(result.file_a.file_path).name,
            file_b_name=Path(result.file_b.file_path).name,
            timestamp=self._format_timestamp(result.timestamp),
            statistics=result.statistics,
            change_percentage=f"{result.statistics.change_percentage * 100:.1f}%",
            rows=rows,
        )

        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding="utf-8")
        return str(output_file)

    def _format_timestamp(self, timestamp: datetime) -> str:
        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=timezone.utc)
        return timestamp.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

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
