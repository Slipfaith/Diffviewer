from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from typing import Iterable

from diff_match_patch import diff_match_patch

from config import FUZZY_MATCH_THRESHOLD, SIMILARITY_THRESHOLD
from core.models import (
    ChangeRecord,
    ChangeStatistics,
    ChangeType,
    ChunkType,
    ComparisonResult,
    DiffChunk,
    ParsedDocument,
    Segment,
)


@dataclass
class MatchResult:
    pairs: list[tuple[Segment, Segment]]
    unmatched_a: list[Segment]
    unmatched_b: list[Segment]


class SegmentMatcher:
    @staticmethod
    def match_by_id(
        segments_a: Iterable[Segment], segments_b: Iterable[Segment]
    ) -> MatchResult:
        list_a = list(segments_a)
        list_b = list(segments_b)
        b_map: dict[str, Segment] = {}
        for segment in list_b:
            if segment.id not in b_map:
                b_map[segment.id] = segment
        matched_ids: set[str] = set()
        pairs: list[tuple[Segment, Segment]] = []
        unmatched_a: list[Segment] = []
        for segment in list_a:
            other = b_map.get(segment.id)
            if other is not None and other.id not in matched_ids:
                pairs.append((segment, other))
                matched_ids.add(other.id)
            else:
                unmatched_a.append(segment)
        unmatched_b = [segment for segment in list_b if segment.id not in matched_ids]
        return MatchResult(pairs=pairs, unmatched_a=unmatched_a, unmatched_b=unmatched_b)

    @staticmethod
    def match_by_position(
        segments_a: Iterable[Segment], segments_b: Iterable[Segment]
    ) -> MatchResult:
        list_a = list(segments_a)
        list_b = list(segments_b)
        pairs = list(zip(list_a, list_b, strict=False))
        min_len = min(len(list_a), len(list_b))
        matched_pairs = [(a, b) for a, b in pairs[:min_len]]
        unmatched_a = list_a[min_len:]
        unmatched_b = list_b[min_len:]
        return MatchResult(
            pairs=matched_pairs, unmatched_a=unmatched_a, unmatched_b=unmatched_b
        )

    @staticmethod
    def match_by_content(
        segments_a: Iterable[Segment],
        segments_b: Iterable[Segment],
        threshold: float = FUZZY_MATCH_THRESHOLD,
    ) -> MatchResult:
        list_a = list(segments_a)
        list_b = list(segments_b)
        used_b: set[int] = set()
        pairs: list[tuple[Segment, Segment]] = []
        unmatched_a: list[Segment] = []

        for segment in list_a:
            best_index: int | None = None
            best_score = threshold
            for idx, candidate in enumerate(list_b):
                if idx in used_b:
                    continue
                score = SequenceMatcher(None, segment.target, candidate.target).ratio()
                if score >= best_score:
                    best_score = score
                    best_index = idx
            if best_index is None:
                unmatched_a.append(segment)
            else:
                used_b.add(best_index)
                pairs.append((segment, list_b[best_index]))

        unmatched_b = [seg for idx, seg in enumerate(list_b) if idx not in used_b]
        return MatchResult(pairs=pairs, unmatched_a=unmatched_a, unmatched_b=unmatched_b)

    @classmethod
    def match(cls, doc_a: ParsedDocument, doc_b: ParsedDocument) -> MatchResult:
        initial = cls.match_by_id(doc_a.segments, doc_b.segments)
        if not initial.unmatched_a and not initial.unmatched_b:
            return initial
        fuzzy = cls.match_by_content(initial.unmatched_a, initial.unmatched_b)
        return MatchResult(
            pairs=initial.pairs + fuzzy.pairs,
            unmatched_a=fuzzy.unmatched_a,
            unmatched_b=fuzzy.unmatched_b,
        )


class TextDiffer:
    _char_threshold = 40

    @staticmethod
    def diff_words(a: str, b: str) -> list[DiffChunk]:
        a_words = a.split()
        b_words = b.split()
        matcher = SequenceMatcher(None, a_words, b_words)
        chunks: list[DiffChunk] = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                text = " ".join(a_words[i1:i2])
                if text:
                    chunks.append(DiffChunk(type=ChunkType.EQUAL, text=text))
            elif tag == "delete":
                text = " ".join(a_words[i1:i2])
                if text:
                    chunks.append(DiffChunk(type=ChunkType.DELETE, text=text))
            elif tag == "insert":
                text = " ".join(b_words[j1:j2])
                if text:
                    chunks.append(DiffChunk(type=ChunkType.INSERT, text=text))
            elif tag == "replace":
                del_text = " ".join(a_words[i1:i2])
                ins_text = " ".join(b_words[j1:j2])
                if del_text:
                    chunks.append(DiffChunk(type=ChunkType.DELETE, text=del_text))
                if ins_text:
                    chunks.append(DiffChunk(type=ChunkType.INSERT, text=ins_text))
        return chunks

    @staticmethod
    def diff_chars(a: str, b: str) -> list[DiffChunk]:
        differ = diff_match_patch()
        diffs = differ.diff_main(a, b)
        differ.diff_cleanupSemantic(diffs)
        chunks: list[DiffChunk] = []
        for op, text in diffs:
            if not text:
                continue
            if op == 0:
                chunks.append(DiffChunk(type=ChunkType.EQUAL, text=text))
            elif op == -1:
                chunks.append(DiffChunk(type=ChunkType.DELETE, text=text))
            elif op == 1:
                chunks.append(DiffChunk(type=ChunkType.INSERT, text=text))
        return chunks

    @classmethod
    def diff_auto(cls, a: str, b: str) -> list[DiffChunk]:
        if max(len(a), len(b)) <= cls._char_threshold:
            return cls.diff_chars(a, b)
        return cls.diff_words(a, b)


class DiffEngine:
    @staticmethod
    def compare(doc_a: ParsedDocument, doc_b: ParsedDocument) -> ComparisonResult:
        match_result = SegmentMatcher.match(doc_a, doc_b)
        changes: list[ChangeRecord] = []

        for seg_a, seg_b in match_result.pairs:
            if seg_a.target == seg_b.target:
                changes.append(
                    ChangeRecord(
                        type=ChangeType.UNCHANGED,
                        segment_before=seg_a,
                        segment_after=seg_b,
                        text_diff=[],
                        similarity=1.0,
                        context=seg_b.context,
                    )
                )
                continue

            similarity = SequenceMatcher(None, seg_a.target, seg_b.target).ratio()
            if similarity >= SIMILARITY_THRESHOLD:
                changes.append(
                    ChangeRecord(
                        type=ChangeType.MODIFIED,
                        segment_before=seg_a,
                        segment_after=seg_b,
                        text_diff=TextDiffer.diff_auto(seg_a.target, seg_b.target),
                        similarity=similarity,
                        context=seg_b.context,
                    )
                )
            else:
                changes.append(
                    ChangeRecord(
                        type=ChangeType.DELETED,
                        segment_before=seg_a,
                        segment_after=None,
                        text_diff=[],
                        similarity=similarity,
                        context=seg_a.context,
                    )
                )
                changes.append(
                    ChangeRecord(
                        type=ChangeType.ADDED,
                        segment_before=None,
                        segment_after=seg_b,
                        text_diff=[],
                        similarity=similarity,
                        context=seg_b.context,
                    )
                )

        for seg_a in match_result.unmatched_a:
            changes.append(
                ChangeRecord(
                    type=ChangeType.DELETED,
                    segment_before=seg_a,
                    segment_after=None,
                    text_diff=[],
                    similarity=0.0,
                    context=seg_a.context,
                )
            )

        for seg_b in match_result.unmatched_b:
            changes.append(
                ChangeRecord(
                    type=ChangeType.ADDED,
                    segment_before=None,
                    segment_after=seg_b,
                    text_diff=[],
                    similarity=0.0,
                    context=seg_b.context,
                )
            )

        statistics = ChangeStatistics.from_changes(changes)
        return ComparisonResult(
            file_a=doc_a,
            file_b=doc_b,
            changes=changes,
            statistics=statistics,
            timestamp=datetime.now(timezone.utc),
        )

    @classmethod
    def compare_multi(cls, docs: list[ParsedDocument]) -> list[ComparisonResult]:
        if len(docs) < 2:
            return []
        results: list[ComparisonResult] = []
        for idx in range(len(docs) - 1):
            results.append(cls.compare(docs[idx], docs[idx + 1]))
        return results
