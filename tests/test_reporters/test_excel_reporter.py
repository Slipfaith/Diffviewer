from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import column_index_from_string

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
from reporters.excel_reporter import ExcelReporter


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


def test_excel_reporter_generates_file(tmp_path: Path) -> None:
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
    ]

    stats = ChangeStatistics.from_changes(changes)
    result = ComparisonResult(
        file_a=make_doc("a.txt", [seg_a, seg_deleted]),
        file_b=make_doc("b.txt", [seg_b, seg_added]),
        changes=changes,
        statistics=stats,
        timestamp=datetime.now(timezone.utc),
    )

    reporter = ExcelReporter()
    output_file = reporter.generate(result, str(tmp_path / "report.xlsx"))
    output_path = Path(output_file)

    assert output_path.exists()
    assert output_path.stat().st_size > 0

    workbook = openpyxl.load_workbook(output_path, read_only=True)
    assert set(workbook.sheetnames) == {"Report", "Statistics"}

    report_ws = workbook["Report"]
    assert report_ws.max_row == len(changes) + 1
    assert report_ws.max_column == 6

    stats_ws = workbook["Statistics"]
    assert stats_ws.max_row >= 5


def test_excel_reporter_column_widths(tmp_path: Path) -> None:
    result = ComparisonResult(
        file_a=make_doc("a.txt", []),
        file_b=make_doc("b.txt", []),
        changes=[],
        statistics=ChangeStatistics.from_changes([]),
        timestamp=datetime.now(timezone.utc),
    )
    reporter = ExcelReporter()
    output_file = reporter.generate(result, str(tmp_path / "widths.xlsx"))
    workbook = openpyxl.load_workbook(output_file)
    ws = workbook["Report"]

    widths = {
        "A": 6,
        "B": 15,
        "C": 30,
        "D": 45,
        "E": 45,
        "F": 12,
    }

    width_map: dict[int, float] = {}
    for dim in ws.column_dimensions.values():
        if dim.min is None or dim.max is None or dim.width is None:
            continue
        for idx in range(dim.min, dim.max + 1):
            width_map[idx] = dim.width

    for col, expected in widths.items():
        idx = column_index_from_string(col)
        assert idx in width_map
        assert width_map[idx] == pytest.approx(expected, abs=1.0)
