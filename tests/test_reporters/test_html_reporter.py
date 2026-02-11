from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from core.diff_engine import DiffEngine
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


def make_segment(segment_id: str, target: str, source: str | None = None) -> Segment:
    context = SegmentContext(
        file_path="file.txt",
        location=segment_id,
        position=int(segment_id),
        group=None,
    )
    return Segment(id=segment_id, source=source, target=target, context=context)


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
                DiffChunk(type=ChunkType.INSERT, text="brave  "),
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
    assert "&middot;" in output_content
    assert "Export to Excel" not in output_content
    assert "function exportToExcel()" not in output_content
    assert "Excel.Sheet" not in output_content
    assert "application/vnd.ms-excel" not in output_content
    assert "alert(" not in output_content
    assert "Added" in output_content
    assert "Deleted" in output_content
    assert "Modified" in output_content
    assert "table-layout: fixed;" in output_content
    assert "word-break: break-word;" in output_content
    assert "overflow-wrap: anywhere;" in output_content
    assert 'class="col-source"' not in output_content
    assert 'class="col-old-target"' in output_content
    assert 'class="col-new-target"' in output_content
    assert ">Type<" not in output_content
    assert "data-type=" not in output_content
    assert 'data-filter="added"' not in output_content
    assert 'data-filter="deleted"' not in output_content
    assert 'data-filter="modified"' not in output_content
    assert 'data-filter="unchanged"' not in output_content


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


def test_html_reporter_shows_source_column_when_available(tmp_path: Path) -> None:
    seg_before = make_segment("1", "Hola", source="Hello")
    seg_after = make_segment("1", "Privet", source="Hello")
    changes = [
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg_before,
            segment_after=seg_after,
            text_diff=[
                DiffChunk(type=ChunkType.DELETE, text="Hola"),
                DiffChunk(type=ChunkType.INSERT, text="Privet"),
            ],
            similarity=0.0,
            context=seg_after.context,
        )
    ]
    result = ComparisonResult(
        file_a=make_doc("a.xliff", [seg_before]),
        file_b=make_doc("b.xliff", [seg_after]),
        changes=changes,
        statistics=ChangeStatistics.from_changes(changes),
        timestamp=datetime.now(timezone.utc),
    )

    reporter = HtmlReporter()
    output_path = Path(reporter.generate(result, str(tmp_path / "with_source.html")))
    output_content = output_path.read_text(encoding="utf-8")
    assert 'class="col-source"' in output_content
    assert ">Source<" in output_content


def test_html_reporter_bidirectional_inline_diff_rules() -> None:
    reporter = HtmlReporter()
    seg_before = make_segment("1", "A old text")
    seg_after = make_segment("1", "A new text")
    modified = ChangeRecord(
        type=ChangeType.MODIFIED,
        segment_before=seg_before,
        segment_after=seg_after,
        text_diff=[
            DiffChunk(type=ChunkType.EQUAL, text="A "),
            DiffChunk(type=ChunkType.DELETE, text="old "),
            DiffChunk(type=ChunkType.INSERT, text="new "),
            DiffChunk(type=ChunkType.EQUAL, text="text"),
        ],
        similarity=0.8,
        context=seg_after.context,
    )

    old_html = reporter._render_old_target(modified)
    new_html = reporter._render_new_target(modified)
    assert "<ins>" not in old_html
    assert "<del>old&middot;</del>" in old_html
    assert "<del>old&middot;</del>" not in new_html
    assert "<ins>new&middot;</ins>" in new_html


def test_html_reporter_added_deleted_rules() -> None:
    reporter = HtmlReporter()
    seg_added = make_segment("1", "Added text")
    seg_deleted = make_segment("2", "Deleted text")

    added = ChangeRecord(
        type=ChangeType.ADDED,
        segment_before=None,
        segment_after=seg_added,
        text_diff=[],
        similarity=0.0,
        context=seg_added.context,
    )
    deleted = ChangeRecord(
        type=ChangeType.DELETED,
        segment_before=seg_deleted,
        segment_after=None,
        text_diff=[],
        similarity=0.0,
        context=seg_deleted.context,
    )

    assert reporter._render_old_target(added) == ""
    assert reporter._render_new_target(added).startswith("<ins>")
    assert reporter._render_new_target(deleted) == ""
    assert reporter._render_old_target(deleted).startswith("<del>")


def test_html_reporter_preserves_newlines(tmp_path: Path) -> None:
    seg_before = make_segment("1", "Line 1\nLine 2")
    seg_after = make_segment("1", "Line 1\nLine 2\nLine 3")
    changes = [
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg_before,
            segment_after=seg_after,
            text_diff=[
                DiffChunk(type=ChunkType.EQUAL, text="Line 1\nLine 2"),
                DiffChunk(type=ChunkType.INSERT, text="\nLine 3"),
            ],
            similarity=0.9,
            context=seg_after.context,
        )
    ]
    result = ComparisonResult(
        file_a=make_doc("a.txt", [seg_before]),
        file_b=make_doc("b.txt", [seg_after]),
        changes=changes,
        statistics=ChangeStatistics.from_changes(changes),
        timestamp=datetime.now(timezone.utc),
    )

    reporter = HtmlReporter()
    output_path = Path(reporter.generate(result, str(tmp_path / "newlines.html")))
    output_content = output_path.read_text(encoding="utf-8")

    assert "white-space: pre-wrap;" in output_content
    assert "Line 1\nLine 2" in output_content


def test_html_reporter_generates_multi_report_with_file_filter(tmp_path: Path) -> None:
    seg_a_before = make_segment("1", "Hello")
    seg_a_after = make_segment("1", "Hello world")
    changes_a = [
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg_a_before,
            segment_after=seg_a_after,
            text_diff=[
                DiffChunk(type=ChunkType.EQUAL, text="Hello"),
                DiffChunk(type=ChunkType.INSERT, text=" world"),
            ],
            similarity=0.9,
            context=seg_a_after.context,
        )
    ]
    result_a = ComparisonResult(
        file_a=make_doc("alpha_a.txt", [seg_a_before]),
        file_b=make_doc("alpha_b.txt", [seg_a_after]),
        changes=changes_a,
        statistics=ChangeStatistics.from_changes(changes_a),
        timestamp=datetime.now(timezone.utc),
    )

    seg_b_before = make_segment("2", "Old")
    seg_b_after = make_segment("2", "New")
    changes_b = [
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg_b_before,
            segment_after=seg_b_after,
            text_diff=[
                DiffChunk(type=ChunkType.DELETE, text="Old"),
                DiffChunk(type=ChunkType.INSERT, text="New"),
            ],
            similarity=0.0,
            context=seg_b_after.context,
        )
    ]
    result_b = ComparisonResult(
        file_a=make_doc("beta_a.txt", [seg_b_before]),
        file_b=make_doc("beta_b.txt", [seg_b_after]),
        changes=changes_b,
        statistics=ChangeStatistics.from_changes(changes_b),
        timestamp=datetime.now(timezone.utc),
    )

    reporter = HtmlReporter()
    output_path = Path(
        reporter.generate_multi(
            [
                ("alpha_a.txt vs alpha_b.txt", result_a),
                ("beta_a.txt vs beta_b.txt", result_b),
            ],
            str(tmp_path / "multi.html"),
        )
    )
    output_content = output_path.read_text(encoding="utf-8")

    assert 'id="file-filter"' in output_content
    assert "alpha_a.txt vs alpha_b.txt" in output_content
    assert "beta_a.txt vs beta_b.txt" in output_content
    assert 'data-file="file-1"' in output_content
    assert 'data-file="file-2"' in output_content
    assert ">File Pair<" in output_content


def test_html_reporter_merges_same_source_change_into_single_row(tmp_path: Path) -> None:
    source_text = "Shared source"
    seg_before = make_segment("100", "Completely old target", source=source_text)
    seg_after = make_segment("200", "Completely new target", source=source_text)
    result = DiffEngine.compare(
        make_doc("a.xliff", [seg_before]),
        make_doc("b.xliff", [seg_after]),
    )

    reporter = HtmlReporter()
    output_path = Path(reporter.generate(result, str(tmp_path / "single_row.html")))
    output_content = output_path.read_text(encoding="utf-8")

    assert output_content.count('class="row-changed') == 1
    assert output_content.count(">Shared source<") == 1
    assert "<del>" in output_content
    assert "<ins>" in output_content
    assert "Completely" in output_content
    assert "target" in output_content
