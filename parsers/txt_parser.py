from __future__ import annotations

from pathlib import Path

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class TxtParser(BaseParser):
    name = "TXT Parser"
    supported_extensions = [".txt"]
    format_description = "Plain Text"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            raw = Path(filepath).read_bytes()
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        encoding = "utf-8"
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            detected = None
            try:
                import chardet

                detected = chardet.detect(raw).get("encoding")
            except Exception:
                detected = None
            encoding = detected or "latin-1"
            text = raw.decode(encoding, errors="replace")

        lines = text.splitlines()
        segments: list[Segment] = []
        for index, line in enumerate(lines, start=1):
            segment_id = str(index)
            context = SegmentContext(
                file_path=filepath,
                location=segment_id,
                position=index,
                group=None,
            )
            segments.append(
                Segment(
                    id=segment_id,
                    source=None,
                    target=line,
                    context=context,
                )
            )

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={},
            encoding=encoding,
        )

    def validate(self, filepath: str) -> list[str]:
        errors: list[str] = []
        try:
            _ = Path(filepath).read_bytes()
        except Exception as exc:
            errors.append(str(exc))
        return errors
