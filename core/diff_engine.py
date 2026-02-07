from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
import re
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
    _token_pattern = re.compile(r"\w+|[^\w\s]|\s", re.UNICODE)
    _word_pattern = re.compile(r"^\w+$", re.UNICODE)

    @classmethod
    def _tokenize(cls, text: str) -> list[str]:
        return cls._token_pattern.findall(text)

    @classmethod
    def _is_word(cls, token: str) -> bool:
        return bool(cls._word_pattern.match(token))

    @staticmethod
    def _append_chunk(chunks: list[DiffChunk], chunk_type: ChunkType, text: str) -> None:
        if not text:
            return
        if chunks and chunks[-1].type == chunk_type:
            chunks[-1].text += text
            return
        chunks.append(DiffChunk(type=chunk_type, text=text))

    @staticmethod
    def diff_words(a: str, b: str) -> list[DiffChunk]:
        a_tokens = TextDiffer._tokenize(a)
        b_tokens = TextDiffer._tokenize(b)
        matcher = SequenceMatcher(None, a_tokens, b_tokens)
        chunks: list[DiffChunk] = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                TextDiffer._append_chunk(
                    chunks, ChunkType.EQUAL, "".join(a_tokens[i1:i2])
                )
            elif tag == "delete":
                TextDiffer._append_chunk(
                    chunks, ChunkType.DELETE, "".join(a_tokens[i1:i2])
                )
            elif tag == "insert":
                TextDiffer._append_chunk(
                    chunks, ChunkType.INSERT, "".join(b_tokens[j1:j2])
                )
            elif tag == "replace":
                TextDiffer._append_replace_chunks(
                    chunks, a_tokens[i1:i2], b_tokens[j1:j2]
                )
        return chunks

    @classmethod
    def _append_replace_chunks(
        cls, chunks: list[DiffChunk], tokens_a: list[str], tokens_b: list[str]
    ) -> None:
        common = min(len(tokens_a), len(tokens_b))
        for idx in range(common):
            token_a = tokens_a[idx]
            token_b = tokens_b[idx]
            if token_a == token_b:
                cls._append_chunk(chunks, ChunkType.EQUAL, token_a)
                continue
            if cls._is_word(token_a) and cls._is_word(token_b):
                if token_a.lower() == token_b.lower():
                    for chunk in cls.diff_chars(token_a, token_b):
                        cls._append_chunk(chunks, chunk.type, chunk.text)
                else:
                    cls._append_chunk(chunks, ChunkType.DELETE, token_a)
                    cls._append_chunk(chunks, ChunkType.INSERT, token_b)
                continue
            cls._append_chunk(chunks, ChunkType.DELETE, token_a)
            cls._append_chunk(chunks, ChunkType.INSERT, token_b)

        for token in tokens_a[common:]:
            cls._append_chunk(chunks, ChunkType.DELETE, token)
        for token in tokens_b[common:]:
            cls._append_chunk(chunks, ChunkType.INSERT, token)

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
        return cls.diff_words(a, b)

    @classmethod
    def has_only_non_word_or_case_changes(cls, a: str, b: str) -> bool:
        words_a = [token.lower() for token in cls._tokenize(a) if cls._is_word(token)]
        words_b = [token.lower() for token in cls._tokenize(b) if cls._is_word(token)]
        return words_a == words_b


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
            text_diff = TextDiffer.diff_auto(seg_a.target, seg_b.target)
            keep_as_modified = similarity >= SIMILARITY_THRESHOLD or TextDiffer.has_only_non_word_or_case_changes(
                seg_a.target, seg_b.target
            )
            if keep_as_modified:
                changes.append(
                    ChangeRecord(
                        type=ChangeType.MODIFIED,
                        segment_before=seg_a,
                        segment_after=seg_b,
                        text_diff=text_diff,
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
