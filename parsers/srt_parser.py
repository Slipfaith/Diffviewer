from __future__ import annotations

from pathlib import Path

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class SrtParser(BaseParser):
    name = "SRT Parser"
    supported_extensions = [".srt"]
    format_description = "SubRip"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            raw = Path(filepath).read_bytes()
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        encoding = "utf-8-sig"
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

        text = text.replace("\r\n", "\n").replace("\r", "\n").strip("\ufeff")
        blocks = [block for block in text.split("\n\n") if block.strip()]
        segments: list[Segment] = []

        for block in blocks:
            lines = [line for line in block.split("\n") if line != ""]
            if len(lines) < 2:
                continue
            subtitle_id = lines[0].strip()
            timecode = lines[1].strip()
            text_lines = lines[2:] if len(lines) > 2 else []
            target = "\n".join(text_lines)

            start, end = "", ""
            if "-->" in timecode:
                parts = [part.strip() for part in timecode.split("-->", 1)]
                if len(parts) == 2:
                    start, end = parts

            context = SegmentContext(
                file_path=filepath,
                location=subtitle_id,
                position=len(segments) + 1,
                group=None,
            )
            segments.append(
                Segment(
                    id=subtitle_id,
                    source=None,
                    target=target,
                    context=context,
                    metadata={"start": start, "end": end},
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
