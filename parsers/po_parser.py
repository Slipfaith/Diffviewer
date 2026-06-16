from __future__ import annotations

import ast
from dataclasses import dataclass, field
import hashlib
from pathlib import Path
import re

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


_FIELD_RE = re.compile(
    r"^(msgctxt|msgid_plural|msgid|msgstr(?:\[(\d+)\])?)\s+(.*)$"
)


@dataclass
class _PoEntry:
    fields: dict[str, str] = field(default_factory=dict)
    references: list[str] = field(default_factory=list)
    flags: list[str] = field(default_factory=list)
    extracted_comments: list[str] = field(default_factory=list)
    translator_comments: list[str] = field(default_factory=list)


class PoParser(BaseParser):
    name = "PO Parser"
    supported_extensions = [".po"]
    format_description = "GNU gettext PO"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            text, encoding = self._read_text(filepath)
            entries = self._parse_entries(text)
        except ParseError:
            raise
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        headers: dict[str, str] = {}
        seen_ids: dict[str, int] = {}

        for entry_index, entry in enumerate(entries, start=1):
            if self._is_header_entry(entry, entry_index):
                headers.update(self._parse_headers(entry.fields.get("msgstr", "")))
                continue
            for segment in self._segments_from_entry(
                entry,
                entry_index=entry_index,
                file_path=filepath,
                seen_ids=seen_ids,
                next_position=lambda: len(segments) + 1,
            ):
                segments.append(segment)

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={"headers": headers} if headers else {},
            encoding=encoding,
        )

    def validate(self, filepath: str) -> list[str]:
        errors: list[str] = []
        try:
            text, _ = self._read_text(filepath)
            self._parse_entries(text)
        except Exception as exc:
            errors.append(str(exc))
        return errors

    @staticmethod
    def _read_text(filepath: str) -> tuple[str, str]:
        try:
            raw = Path(filepath).read_bytes()
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        encoding = "utf-8-sig"
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            detected = None
            try:
                import chardet

                detected = chardet.detect(raw).get("encoding")
            except Exception:
                detected = None
            encoding = detected or "latin-1"
            return raw.decode(encoding, errors="replace"), encoding

    @classmethod
    def _parse_entries(cls, text: str) -> list[_PoEntry]:
        entries: list[_PoEntry] = []
        entry = _PoEntry()
        current_field: str | None = None

        def flush() -> None:
            nonlocal entry, current_field
            if entry.fields:
                entries.append(entry)
            entry = _PoEntry()
            current_field = None

        for line_no, line in enumerate(text.splitlines(), start=1):
            stripped = line.strip()
            if not stripped:
                flush()
                continue

            if stripped.startswith("#~"):
                current_field = None
                continue

            if stripped.startswith("#"):
                cls._capture_comment(entry, stripped)
                current_field = None
                continue

            field_match = _FIELD_RE.match(stripped)
            if field_match is not None:
                field_name = field_match.group(1)
                entry.fields[field_name] = cls._decode_quoted(
                    field_match.group(3),
                    line_no,
                )
                current_field = field_name
                continue

            if stripped.startswith('"'):
                if current_field is None:
                    raise ValueError(f"line {line_no}: quoted string without a field")
                entry.fields[current_field] += cls._decode_quoted(stripped, line_no)
                continue

            raise ValueError(f"line {line_no}: expected PO field or quoted string")

        flush()
        return entries

    @staticmethod
    def _decode_quoted(raw_value: str, line_no: int) -> str:
        value = raw_value.strip()
        if not value.startswith('"'):
            raise ValueError(f"line {line_no}: expected quoted string")
        try:
            decoded = ast.literal_eval(value)
        except Exception as exc:
            raise ValueError(f"line {line_no}: invalid quoted string") from exc
        if not isinstance(decoded, str):
            raise ValueError(f"line {line_no}: expected quoted string")
        return decoded

    @staticmethod
    def _capture_comment(entry: _PoEntry, stripped_line: str) -> None:
        if stripped_line.startswith("#."):
            text = stripped_line[2:].strip()
            if text:
                entry.extracted_comments.append(text)
            return
        if stripped_line.startswith("#:"):
            entry.references.extend(stripped_line[2:].split())
            return
        if stripped_line.startswith("#,"):
            flags = [flag.strip() for flag in stripped_line[2:].split(",")]
            entry.flags.extend(flag for flag in flags if flag)
            return
        if stripped_line.startswith("# "):
            text = stripped_line[2:].strip()
            if text:
                entry.translator_comments.append(text)

    @staticmethod
    def _is_header_entry(entry: _PoEntry, entry_index: int) -> bool:
        return (
            entry_index == 1
            and entry.fields.get("msgid", "") == ""
            and "msgstr" in entry.fields
            and "msgctxt" not in entry.fields
            and "msgid_plural" not in entry.fields
        )

    @staticmethod
    def _parse_headers(raw_header: str) -> dict[str, str]:
        headers: dict[str, str] = {}
        for line in raw_header.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            key = key.strip()
            if key:
                headers[key] = value.strip()
        return headers

    @classmethod
    def _segments_from_entry(
        cls,
        entry: _PoEntry,
        *,
        entry_index: int,
        file_path: str,
        seen_ids: dict[str, int],
        next_position,
    ) -> list[Segment]:
        source = entry.fields.get("msgid", "")
        plural_source = entry.fields.get("msgid_plural")
        plural_targets = cls._plural_targets(entry)

        if plural_targets:
            segments: list[Segment] = []
            for plural_index, target in plural_targets:
                segment_source = (
                    source if plural_index == "0" else plural_source or source
                )
                metadata = cls._metadata_for_entry(entry)
                metadata["plural_index"] = plural_index
                if plural_source is not None:
                    metadata["plural_source"] = plural_source
                segments.append(
                    cls._make_segment(
                        file_path=file_path,
                        entry_index=entry_index,
                        position=next_position() + len(segments),
                        source=segment_source,
                        target=target,
                        metadata=metadata,
                        seen_ids=seen_ids,
                        key_suffix=f"plural:{plural_index}",
                    )
                )
            return segments

        return [
            cls._make_segment(
                file_path=file_path,
                entry_index=entry_index,
                position=next_position(),
                source=source,
                target=entry.fields.get("msgstr", ""),
                metadata=cls._metadata_for_entry(entry),
                seen_ids=seen_ids,
                key_suffix="msgstr",
            )
        ]

    @staticmethod
    def _plural_targets(entry: _PoEntry) -> list[tuple[str, str]]:
        values: list[tuple[int, str, str]] = []
        for field_name, target in entry.fields.items():
            match = re.fullmatch(r"msgstr\[(\d+)\]", field_name)
            if match is not None:
                values.append((int(match.group(1)), match.group(1), target))
        return [(raw_index, target) for _, raw_index, target in sorted(values)]

    @classmethod
    def _make_segment(
        cls,
        *,
        file_path: str,
        entry_index: int,
        position: int,
        source: str,
        target: str,
        metadata: dict[str, object],
        seen_ids: dict[str, int],
        key_suffix: str,
    ) -> Segment:
        segment_id = cls._stable_segment_id(
            metadata.get("context", ""),
            source,
            key_suffix,
        )
        seen_count = seen_ids.get(segment_id, 0) + 1
        seen_ids[segment_id] = seen_count
        if seen_count > 1:
            segment_id = f"{segment_id}#{seen_count}"

        context = SegmentContext(
            file_path=file_path,
            location=f"entry {entry_index}",
            position=position,
            group=(
                metadata.get("context")
                if isinstance(metadata.get("context"), str)
                else None
            ),
        )
        return Segment(
            id=segment_id,
            source=source,
            target=target,
            context=context,
            metadata=metadata,
        )

    @staticmethod
    def _stable_segment_id(*parts: object) -> str:
        key = "\x04".join(str(part or "") for part in parts)
        digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:12]
        return f"po:{digest}"

    @staticmethod
    def _metadata_for_entry(entry: _PoEntry) -> dict[str, object]:
        metadata: dict[str, object] = {}
        context = entry.fields.get("msgctxt")
        if context is not None:
            metadata["context"] = context
        if entry.references:
            metadata["references"] = list(entry.references)
        if entry.flags:
            metadata["flags"] = list(entry.flags)
        if entry.extracted_comments:
            metadata["extracted_comments"] = list(entry.extracted_comments)
        if entry.translator_comments:
            metadata["translator_comments"] = list(entry.translator_comments)
        return metadata
