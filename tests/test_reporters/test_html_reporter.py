from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from core.models import (
    ChangeRecord,
    ChangeStatistics,
    ChangeType,
    ChunkType,
    ComparisonResult,
    DiffChunk,
    ParsedDocument,
    Segment,
    SegmentContext,
)
from reporters.html_reporter import HtmlReporter


def make_segment(segment_id: str, target: str) -> Segment:
    context = SegmentContext(
        file_path="file.txt",
        location=segment_id,
        position=int(segment_id),
        group=None,
    )
    return Segment(id=segment_id, source=None, target=target, context=context)


def make_doc(name: str, segments: list[Segment]) -> ParsedDocument:
    return ParsedDocument(segments=segments, format_name="TXT", file_path=name)


def test_html_reporter_generates_report(tmp_path: Path) -> None:
    seg_a = make_segment("1", "Hello world")
    seg_b = make_segment("1", "Hello brave world")
    seg_added = make_segment("2", "New line")
    seg_deleted = make_segment("3", "Old line")

    changes = [
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg_a,
            segment_after=seg_b,
            text_diff=[
                DiffChunk(type=ChunkType.EQUAL, text="Hello "),
                DiffChunk(type=ChunkType.DELETE, text="old "),
                DiffChunk(type=ChunkType.INSERT, text="brave "),
                DiffChunk(type=ChunkType.EQUAL, text="world"),
            ],
            similarity=0.8,
            context=seg_b.context,
        ),
        ChangeRecord(
            type=ChangeType.ADDED,
            segment_before=None,
            segment_after=seg_added,
            text_diff=[],
            similarity=0.0,
            context=seg_added.context,
        ),
        ChangeRecord(
            type=ChangeType.DELETED,
            segment_before=seg_deleted,
            segment_after=None,
            text_diff=[],
            similarity=0.0,
            context=seg_deleted.context,
        ),
        ChangeRecord(
            type=ChangeType.UNCHANGED,
            segment_before=seg_a,
            segment_after=seg_a,
            text_diff=[],
            similarity=1.0,
            context=seg_a.context,
        ),
    ]

    stats = ChangeStatistics.from_changes(changes)
    result = ComparisonResult(
        file_a=make_doc("a.txt", [seg_a, seg_deleted]),
        file_b=make_doc("b.txt", [seg_b, seg_added]),
        changes=changes,
        statistics=stats,
        timestamp=datetime.now(timezone.utc),
    )

    reporter = HtmlReporter()
    output_path = tmp_path / "report.html"
    output_file = reporter.generate(result, str(output_path))

    output_content = Path(output_file).read_text(encoding="utf-8")
    assert "<html" in output_content
    assert "a.txt" in output_content
    assert "b.txt" in output_content
    assert "<ins>" in output_content
    assert "<del>" in output_content
    assert "Added" in output_content
    assert "Deleted" in output_content
    assert "Modified" in output_content


def test_html_reporter_empty_result(tmp_path: Path) -> None:
    doc = make_doc("a.txt", [])
    result = ComparisonResult(
        file_a=doc,
        file_b=doc,
        changes=[],
        statistics=ChangeStatistics.from_changes([]),
        timestamp=datetime.now(timezone.utc),
    )
    reporter = HtmlReporter()
    output_file = reporter.generate(result, str(tmp_path / "empty.html"))
    assert Path(output_file).exists()
