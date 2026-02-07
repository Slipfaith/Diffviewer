from __future__ import annotations

from pathlib import Path

from docx import Document

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class DocxParser(BaseParser):
    name = "DOCX Parser"
    supported_extensions = [".docx", ".doc"]
    format_description = "Word Document"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        ext = Path(filepath).suffix.lower()
        if ext == ".doc":
            raise NotImplementedError("DOC format requires conversion to DOCX first")

        try:
            document = Document(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        for index, paragraph in enumerate(document.paragraphs, start=1):
            text = paragraph.text
            if not text or text.strip() == "":
                continue
            segment_id = f"para_{index}"
            context = SegmentContext(
                file_path=filepath,
                location=f"Paragraph {index}",
                position=len(segments) + 1,
                group=None,
            )
            style_name = paragraph.style.name if paragraph.style is not None else ""
            segments.append(
                Segment(
                    id=segment_id,
                    source=None,
                    target=text,
                    context=context,
                    metadata={"style": style_name},
                )
            )

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={},
            encoding=None,
        )

    def validate(self, filepath: str) -> list[str]:
        errors: list[str] = []
        try:
            ext = Path(filepath).suffix.lower()
            if ext == ".doc":
                errors.append("DOC format requires conversion to DOCX first")
                return errors
            _ = Document(filepath)
        except Exception as exc:
            errors.append(str(exc))
        return errors
