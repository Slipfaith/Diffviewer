from datetime import datetime, timezone

import pytest

from core.models import (
    ChangeRecord,
    ChangeStatistics,
    ChangeType,
    ChunkType,
    ComparisonError,
    ComparisonResult,
    DiffChunk,
    FormatRun,
    ParseError,
    ParsedDocument,
    Segment,
    SegmentContext,
    UnsupportedFormatError,
)


def make_context() -> SegmentContext:
    return SegmentContext(file_path="a.txt", location="line 1", position=1, group=None)


def make_segment(segment_id: str, source: str | None, target: str) -> Segment:
    return Segment(
        id=segment_id,
        source=source,
        target=target,
        context=make_context(),
        formatting=[],
        metadata={},
    )


def test_segment_properties() -> None:
    segment = make_segment("1", "src", "tgt")
    assert segment.plain_text == "tgt"
    assert segment.has_source is True

    segment_no_source = make_segment("2", None, "tgt")
    assert segment_no_source.has_source is False

    segment_empty_source = make_segment("3", "", "tgt")
    assert segment_empty_source.has_source is False


def test_format_run_and_diff_chunk_creation() -> None:
    run = FormatRun(text="hi", bold=True, color="#ff0000")
    chunk = DiffChunk(type=ChunkType.INSERT, text="hi", formatting=[run])
    assert chunk.type == ChunkType.INSERT
    assert chunk.text == "hi"
    assert chunk.formatting == [run]


def test_parsed_document_helpers() -> None:
    seg_a = make_segment("a", None, "one")
    seg_b = make_segment("b", None, "two")
    doc = ParsedDocument(segments=[seg_a, seg_b], format_name="TXT", file_path="a.txt")

    assert doc.segment_count == 2
    assert doc.get_segment_by_id("a") is seg_a

    with pytest.raises(KeyError):
        doc.get_segment_by_id("missing")


def test_change_record_is_changed() -> None:
    ctx = make_context()
    seg = make_segment("1", None, "tgt")
    change = ChangeRecord(
        type=ChangeType.MODIFIED,
        segment_before=seg,
        segment_after=seg,
        text_diff=[],
        similarity=0.5,
        context=ctx,
    )
    assert change.is_changed is True

    unchanged = ChangeRecord(
        type=ChangeType.UNCHANGED,
        segment_before=seg,
        segment_after=seg,
        text_diff=[],
        similarity=1.0,
        context=ctx,
    )
    assert unchanged.is_changed is False


def test_change_statistics_from_changes() -> None:
    ctx = make_context()
    seg = make_segment("1", None, "tgt")
    changes = [
        ChangeRecord(
            type=ChangeType.ADDED,
            segment_before=None,
            segment_after=seg,
            text_diff=[],
            similarity=0.0,
            context=ctx,
        ),
        ChangeRecord(
            type=ChangeType.DELETED,
            segment_before=seg,
            segment_after=None,
            text_diff=[],
            similarity=0.0,
            context=ctx,
        ),
        ChangeRecord(
            type=ChangeType.MODIFIED,
            segment_before=seg,
            segment_after=seg,
            text_diff=[],
            similarity=0.5,
            context=ctx,
        ),
        ChangeRecord(
            type=ChangeType.MOVED,
            segment_before=seg,
            segment_after=seg,
            text_diff=[],
            similarity=0.8,
            context=ctx,
        ),
        ChangeRecord(
            type=ChangeType.UNCHANGED,
            segment_before=seg,
            segment_after=seg,
            text_diff=[],
            similarity=1.0,
            context=ctx,
        ),
    ]
    stats = ChangeStatistics.from_changes(changes)
    assert stats.total_segments == 5
    assert stats.added == 1
    assert stats.deleted == 1
    assert stats.modified == 1
    assert stats.moved == 1
    assert stats.unchanged == 1
    assert stats.change_percentage == pytest.approx(4 / 5)


def test_change_statistics_zero_division() -> None:
    stats = ChangeStatistics.from_changes([])
    assert stats.total_segments == 0
    assert stats.change_percentage == 0.0


def test_comparison_result_change_percentage() -> None:
    seg = make_segment("1", None, "tgt")
    doc = ParsedDocument(segments=[seg], format_name="TXT", file_path="a.txt")
    stats = ChangeStatistics.from_changes([])
    result = ComparisonResult(
        file_a=doc,
        file_b=doc,
        changes=[],
        statistics=stats,
        timestamp=datetime.now(timezone.utc),
    )
    assert result.change_percentage == stats.change_percentage


def test_custom_exceptions() -> None:
    parse_error = ParseError("file.txt", "bad encoding")
    assert parse_error.filepath == "file.txt"
    assert parse_error.reason == "bad encoding"

    unsupported = UnsupportedFormatError(".abc")
    assert unsupported.extension == ".abc"

    assert issubclass(ComparisonError, Exception)
