from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from collections import deque
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
    def match(
        cls,
        doc_a: ParsedDocument,
        doc_b: ParsedDocument,
        allow_fuzzy: bool = True,
    ) -> MatchResult:
        initial = cls.match_by_id(doc_a.segments, doc_b.segments)
        if (
            not allow_fuzzy
            or (not initial.unmatched_a and not initial.unmatched_b)
        ):
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
        if "\n" in a or "\n" in b:
            return cls._diff_lines_then_words(a, b)
        return cls.diff_words(a, b)

    @classmethod
    def _diff_lines_then_words(cls, a: str, b: str) -> list[DiffChunk]:
        a_norm = a.replace("\r\n", "\n").replace("\r", "\n")
        b_norm = b.replace("\r\n", "\n").replace("\r", "\n")
        a_lines = a_norm.splitlines(keepends=True)
        b_lines = b_norm.splitlines(keepends=True)
        # Ensure last element ends with \n for consistent matching
        if a_lines and not a_lines[-1].endswith("\n"):
            a_lines[-1] += "\n"
        if b_lines and not b_lines[-1].endswith("\n"):
            b_lines[-1] += "\n"

        matcher = SequenceMatcher(None, a_lines, b_lines)
        chunks: list[DiffChunk] = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                cls._append_chunk(chunks, ChunkType.EQUAL, "".join(a_lines[i1:i2]))
            elif tag == "delete":
                cls._append_chunk(chunks, ChunkType.DELETE, "".join(a_lines[i1:i2]))
            elif tag == "insert":
                cls._append_chunk(chunks, ChunkType.INSERT, "".join(b_lines[j1:j2]))
            elif tag == "replace":
                cls._diff_replace_lines(chunks, a_lines[i1:i2], b_lines[j1:j2])
        return chunks

    @classmethod
    def _diff_replace_lines(
        cls,
        chunks: list[DiffChunk],
        old_lines: list[str],
        new_lines: list[str],
    ) -> None:
        common = min(len(old_lines), len(new_lines))
        for i in range(common):
            line_chunks = cls.diff_words(old_lines[i], new_lines[i])
            for chunk in line_chunks:
                cls._append_chunk(chunks, chunk.type, chunk.text)
        for line in old_lines[common:]:
            cls._append_chunk(chunks, ChunkType.DELETE, line)
        for line in new_lines[common:]:
            cls._append_chunk(chunks, ChunkType.INSERT, line)

    @classmethod
    def has_only_non_word_or_case_changes(cls, a: str, b: str) -> bool:
        words_a = [token.lower() for token in cls._tokenize(a) if cls._is_word(token)]
        words_b = [token.lower() for token in cls._tokenize(b) if cls._is_word(token)]
        return words_a == words_b


class DiffEngine:
    @staticmethod
    def _normalize_source(source: str | None) -> str:
        return " ".join((source or "").casefold().split())

    @classmethod
    def _sources_match(cls, source_a: str | None, source_b: str | None) -> bool:
        normalized_a = cls._normalize_source(source_a)
        normalized_b = cls._normalize_source(source_b)
        return bool(normalized_a and normalized_b and normalized_a == normalized_b)

    @classmethod
    def _pair_unmatched_by_source(cls, match_result: MatchResult) -> MatchResult:
        if not match_result.unmatched_a or not match_result.unmatched_b:
            return match_result

        by_source_b: dict[str, deque[Segment]] = {}
        for segment in match_result.unmatched_b:
            source_key = cls._normalize_source(segment.source)
            if source_key:
                by_source_b.setdefault(source_key, deque()).append(segment)

        source_pairs: list[tuple[Segment, Segment]] = []
        for segment_a in match_result.unmatched_a:
            source_key = cls._normalize_source(segment_a.source)
            if not source_key:
                continue
            candidates = by_source_b.get(source_key)
            if not candidates:
                continue
            source_pairs.append((segment_a, candidates.popleft()))

        if not source_pairs:
            return match_result

        paired_a_ids = {id(segment_a) for segment_a, _ in source_pairs}
        paired_b_ids = {id(segment_b) for _, segment_b in source_pairs}
        unmatched_a = [
            segment for segment in match_result.unmatched_a if id(segment) not in paired_a_ids
        ]
        unmatched_b = [
            segment for segment in match_result.unmatched_b if id(segment) not in paired_b_ids
        ]
        return MatchResult(
            pairs=match_result.pairs + source_pairs,
            unmatched_a=unmatched_a,
            unmatched_b=unmatched_b,
        )

    @staticmethod
    def _is_xliff_family(doc_a: ParsedDocument, doc_b: ParsedDocument) -> bool:
        format_a = (doc_a.format_name or "").upper()
        format_b = (doc_b.format_name or "").upper()
        return "XLIFF" in format_a and "XLIFF" in format_b

    @staticmethod
    def _is_sdlxliff(doc_a: ParsedDocument, doc_b: ParsedDocument) -> bool:
        return (
            (doc_a.format_name or "").upper() == "SDLXLIFF"
            and (doc_b.format_name or "").upper() == "SDLXLIFF"
        )

    @staticmethod
    def compare(doc_a: ParsedDocument, doc_b: ParsedDocument) -> ComparisonResult:
        strict_id_mode = DiffEngine._is_sdlxliff(doc_a, doc_b)
        xliff_family_mode = DiffEngine._is_xliff_family(doc_a, doc_b)
        match_result = SegmentMatcher.match(
            doc_a, doc_b, allow_fuzzy=not strict_id_mode
        )
        match_result = DiffEngine._pair_unmatched_by_source(match_result)
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
            ids_match = seg_a.id == seg_b.id
            keep_as_modified = (
                strict_id_mode
                or ids_match
                or DiffEngine._sources_match(seg_a.source, seg_b.source)
                or similarity >= SIMILARITY_THRESHOLD
                or TextDiffer.has_only_non_word_or_case_changes(
                    seg_a.target, seg_b.target
                )
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
