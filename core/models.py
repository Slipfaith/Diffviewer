from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any, Iterable


class ParseError(Exception):
    def __init__(self, filepath: str, reason: str) -> None:
        super().__init__(f"Failed to parse '{filepath}': {reason}")
        self.filepath = filepath
        self.reason = reason


class UnsupportedFormatError(Exception):
    def __init__(self, extension: str) -> None:
        super().__init__(f"Unsupported format: {extension}")
        self.extension = extension


class ComparisonError(Exception):
    pass


class ChangeType(str, Enum):
    ADDED = "ADDED"
    DELETED = "DELETED"
    MODIFIED = "MODIFIED"
    MOVED = "MOVED"
    UNCHANGED = "UNCHANGED"


class ChunkType(str, Enum):
    EQUAL = "EQUAL"
    INSERT = "INSERT"
    DELETE = "DELETE"


@dataclass
class SegmentContext:
    file_path: str
    location: str
    position: int
    group: str | None = None


@dataclass
class FormatRun:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    color: str | None = None
    font: str | None = None
    size: float | None = None


@dataclass
class Segment:
    id: str
    source: str | None
    target: str
    context: SegmentContext
    formatting: list[FormatRun] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def plain_text(self) -> str:
        return self.target

    @property
    def has_source(self) -> bool:
        return bool(self.source)


@dataclass
class ParsedDocument:
    segments: list[Segment]
    format_name: str
    file_path: str
    metadata: dict[str, Any] = field(default_factory=dict)
    encoding: str | None = None

    @property
    def segment_count(self) -> int:
        return len(self.segments)

    def get_segment_by_id(self, segment_id: str) -> Segment:
        for segment in self.segments:
            if segment.id == segment_id:
                return segment
        raise KeyError(f"Segment not found: {segment_id}")


@dataclass
class DiffChunk:
    type: ChunkType
    text: str
    formatting: list[FormatRun] | None = None


@dataclass
class ChangeRecord:
    type: ChangeType
    segment_before: Segment | None
    segment_after: Segment | None
    text_diff: list[DiffChunk]
    similarity: float
    context: SegmentContext

    @property
    def is_changed(self) -> bool:
        return self.type != ChangeType.UNCHANGED


@dataclass
class ChangeStatistics:
    total_segments: int
    added: int
    deleted: int
    modified: int
    moved: int
    unchanged: int
    change_percentage: float

    @classmethod
    def from_changes(cls, changes: Iterable[ChangeRecord]) -> "ChangeStatistics":
        items = list(changes)
        total = len(items)
        added = deleted = modified = moved = unchanged = 0
        for change in items:
            if change.type == ChangeType.ADDED:
                added += 1
            elif change.type == ChangeType.DELETED:
                deleted += 1
            elif change.type == ChangeType.MODIFIED:
                modified += 1
            elif change.type == ChangeType.MOVED:
                moved += 1
            elif change.type == ChangeType.UNCHANGED:
                unchanged += 1
        changed = added + deleted + modified + moved
        percentage = 0.0 if total == 0 else changed / total
        return cls(
            total_segments=total,
            added=added,
            deleted=deleted,
            modified=modified,
            moved=moved,
            unchanged=unchanged,
            change_percentage=percentage,
        )


@dataclass
class ComparisonResult:
    file_a: ParsedDocument
    file_b: ParsedDocument
    changes: list[ChangeRecord]
    statistics: ChangeStatistics
    timestamp: datetime

    @property
    def change_percentage(self) -> float:
        return self.statistics.change_percentage


@dataclass
class BatchFileResult:
    filename: str
    status: str
    report_paths: list[str] = field(default_factory=list)
    statistics: ChangeStatistics | None = None
    comparison: ComparisonResult | None = None
    error_message: str | None = None


@dataclass
class BatchResult:
    folder_a: str
    folder_b: str
    files: list[BatchFileResult]
    summary_report_path: str | None = None
    summary_excel_path: str | None = None

    @property
    def total_files(self) -> int:
        return len(self.files)

    @property
    def compared_files(self) -> int:
        return sum(1 for item in self.files if item.status == "compared")

    @property
    def only_in_a(self) -> int:
        return sum(1 for item in self.files if item.status == "only_in_a")

    @property
    def only_in_b(self) -> int:
        return sum(1 for item in self.files if item.status == "only_in_b")

    @property
    def errors(self) -> int:
        return sum(1 for item in self.files if item.status == "error")


@dataclass
class MultiVersionResult:
    file_paths: list[str]
    comparisons: list[ComparisonResult]
    documents: list[ParsedDocument] = field(default_factory=list)
    report_paths: list[list[str]] = field(default_factory=list)
    summary_report_path: str | None = None
