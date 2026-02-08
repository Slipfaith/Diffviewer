from __future__ import annotations

import re
from pathlib import Path

from core.models import MultiVersionResult, ParsedDocument, Segment, SegmentContext
from reporters.summary_reporter import SummaryReporter


def _segment(segment_id: str, source: str, target: str, position: int) -> Segment:
    return Segment(
        id=segment_id,
        source=source,
        target=target,
        context=SegmentContext(
            file_path="test.xliff",
            location=segment_id,
            position=position,
            group=None,
        ),
    )


def _doc(path: str, segments: list[Segment]) -> ParsedDocument:
    return ParsedDocument(segments=segments, format_name="XLIFF", file_path=path)


def test_summary_reporter_marks_non_text_changes_symbol_level(tmp_path: Path) -> None:
    doc1 = _doc("v1.xliff", [_segment("1", "Greeting", "Hello, world", 1)])
    doc2 = _doc("v2.xliff", [_segment("1", "Greeting", "Hello world", 1)])
    doc3 = _doc("v3.xliff", [_segment("1", "Greeting", "Hello  world", 1)])

    result = MultiVersionResult(
        file_paths=["v1.xliff", "v2.xliff", "v3.xliff"],
        comparisons=[],
        documents=[doc1, doc2, doc3],
    )
    output = SummaryReporter().generate_versions(result, str(tmp_path / "versions_summary.html"))
    text = Path(output).read_text(encoding="utf-8")

    assert "version-del-1" in text
    assert "symbol-del" in text
    assert "ws-change" in text
    assert "data-filter=\"changed\"" in text


def test_summary_reporter_fills_targets_by_source_when_ids_change(tmp_path: Path) -> None:
    doc1 = _doc(
        "v1.sdlxliff",
        [
            _segment("1", "Source one.", "Target one v1", 1),
            _segment("2", "Source two?", "Target two v1", 2),
        ],
    )
    doc2 = _doc(
        "v2.sdlxliff",
        [
            _segment("A-101", "Source one.", "Target one v2", 1),
            _segment("A-202", "Source two?", "Target two v2", 2),
        ],
    )
    doc3 = _doc(
        "v3.sdlxliff",
        [
            _segment("B-501", "Source one", "Target one v3", 1),
            _segment("B-502", "Source two", "Target two v3", 2),
        ],
    )

    result = MultiVersionResult(
        file_paths=["v1.sdlxliff", "v2.sdlxliff", "v3.sdlxliff"],
        comparisons=[],
        documents=[doc1, doc2, doc3],
    )
    output = SummaryReporter().generate_versions(result, str(tmp_path / "versions_summary.html"))
    text = Path(output).read_text(encoding="utf-8")
    plain = re.sub(r"<[^>]+>", "", text)

    assert "Target one v2" in plain
    assert "Target one v3" in plain
    assert "Target two v2" in plain
    assert "Target two v3" in plain
    assert "<td class=\"state-missing\"></td>" not in text
