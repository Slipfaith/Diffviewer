from __future__ import annotations

from core.diff_engine import DiffEngine
from core.models import ChangeType, ChunkType, ParsedDocument, Segment, SegmentContext


def make_segment(segment_id: str, target: str, position: int) -> Segment:
    context = SegmentContext(
        file_path="file.txt",
        location=f"line {position}",
        position=position,
        group=None,
    )
    return Segment(id=segment_id, source=None, target=target, context=context)


def make_doc(segments: list[Segment], name: str = "DOC") -> ParsedDocument:
    return ParsedDocument(segments=segments, format_name=name, file_path="file.txt")


def test_identical_documents_unchanged() -> None:
    segs = [make_segment("1", "hello world", 1), make_segment("2", "second", 2)]
    doc_a = make_doc(segs)
    doc_b = make_doc([make_segment("1", "hello world", 1), make_segment("2", "second", 2)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 2
    assert all(change.type == ChangeType.UNCHANGED for change in result.changes)
    assert result.statistics.unchanged == 2
    assert result.statistics.change_percentage == 0.0


def test_added_segment() -> None:
    doc_a = make_doc([make_segment("1", "hello", 1)])
    doc_b = make_doc([make_segment("1", "hello", 1), make_segment("2", "new", 2)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert result.statistics.added == 1
    assert result.statistics.deleted == 0
    assert any(change.type == ChangeType.ADDED for change in result.changes)


def test_deleted_segment() -> None:
    doc_a = make_doc([make_segment("1", "hello", 1), make_segment("2", "old", 2)])
    doc_b = make_doc([make_segment("1", "hello", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert result.statistics.deleted == 1
    assert result.statistics.added == 0
    assert any(change.type == ChangeType.DELETED for change in result.changes)


def test_modified_segment_with_text_diff() -> None:
    doc_a = make_doc([make_segment("1", "hello world", 1)])
    doc_b = make_doc([make_segment("1", "hello brave world", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert change.text_diff
    assert any(chunk.type != ChunkType.EQUAL for chunk in change.text_diff)


def test_similarity_below_threshold_becomes_added_deleted() -> None:
    doc_a = make_doc([make_segment("1", "aaa", 1)])
    doc_b = make_doc([make_segment("1", "bbbbbbbbbbbbbbbbbbbb", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert result.statistics.added == 1
    assert result.statistics.deleted == 1
    assert result.statistics.modified == 0
    assert all(change.type != ChangeType.MODIFIED for change in result.changes)


def test_empty_documents() -> None:
    doc_a = make_doc([])
    doc_b = make_doc([])

    result = DiffEngine.compare(doc_a, doc_b)
    assert result.changes == []
    assert result.statistics.total_segments == 0
    assert result.statistics.change_percentage == 0.0


def test_compare_multi_chain() -> None:
    doc_a = make_doc([make_segment("1", "one", 1)], name="A")
    doc_b = make_doc([make_segment("1", "two", 1)], name="B")
    doc_c = make_doc([make_segment("1", "three", 1)], name="C")

    results = DiffEngine.compare_multi([doc_a, doc_b, doc_c])
    assert len(results) == 2
    assert results[0].file_a is doc_a
    assert results[0].file_b is doc_b
    assert results[1].file_a is doc_b
    assert results[1].file_b is doc_c


def test_modified_when_only_space_changed() -> None:
    doc_a = make_doc([make_segment("1", "Hello world", 1)])
    doc_b = make_doc([make_segment("1", "Hello  world", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert any(chunk.type == ChunkType.INSERT and chunk.text == " " for chunk in change.text_diff)


def test_modified_when_only_punctuation_changed() -> None:
    doc_a = make_doc([make_segment("1", "Hello, world", 1)])
    doc_b = make_doc([make_segment("1", "Hello. world", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert any(chunk.type == ChunkType.DELETE and chunk.text == "," for chunk in change.text_diff)
    assert any(chunk.type == ChunkType.INSERT and chunk.text == "." for chunk in change.text_diff)


def test_word_replacement_is_whole_word_delete_insert() -> None:
    doc_a = make_doc([make_segment("1", "молоко", 1)])
    doc_b = make_doc([make_segment("1", "болото", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert len(change.text_diff) == 2
    assert change.text_diff[0].type == ChunkType.DELETE
    assert change.text_diff[0].text == "молоко"
    assert change.text_diff[1].type == ChunkType.INSERT
    assert change.text_diff[1].text == "болото"


def test_case_only_change_uses_char_level_diff() -> None:
    doc_a = make_doc([make_segment("1", "Hello", 1)])
    doc_b = make_doc([make_segment("1", "hello", 1)])

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert any(chunk.type == ChunkType.DELETE and chunk.text == "H" for chunk in change.text_diff)
    assert any(chunk.type == ChunkType.INSERT and chunk.text == "h" for chunk in change.text_diff)


def test_sdlxliff_same_id_low_similarity_stays_modified() -> None:
    doc_a = make_doc([make_segment("209", "aaa", 1)], name="SDLXLIFF")
    doc_b = make_doc([make_segment("209", "bbbbbbbbbbbbbbbbbbbb", 1)], name="SDLXLIFF")

    result = DiffEngine.compare(doc_a, doc_b)
    assert len(result.changes) == 1
    assert result.statistics.modified == 1
    assert result.statistics.added == 0
    assert result.statistics.deleted == 0
    change = result.changes[0]
    assert change.type == ChangeType.MODIFIED
    assert change.segment_before is not None
    assert change.segment_after is not None
    assert change.segment_before.id == "209"
    assert change.segment_after.id == "209"


def test_sdlxliff_does_not_fuzzy_match_different_ids() -> None:
    doc_a = make_doc([make_segment("101", "Same text", 1)], name="SDLXLIFF")
    doc_b = make_doc([make_segment("202", "Same text", 1)], name="SDLXLIFF")

    result = DiffEngine.compare(doc_a, doc_b)
    assert result.statistics.modified == 0
    assert result.statistics.unchanged == 0
    assert result.statistics.added == 1
    assert result.statistics.deleted == 1
    assert all(change.type != ChangeType.MODIFIED for change in result.changes)
