from __future__ import annotations

from pathlib import Path

from docx import Document

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class DocxParser(BaseParser):
    name = "DOCX Parser"
    supported_extensions = [".docx"]
    format_description = "Word Document"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() == ".docx"

    def parse(self, filepath: str) -> ParsedDocument:
        ext = Path(filepath).suffix.lower()
        if ext == ".doc":
            raise NotImplementedError("DOC format requires conversion to DOCX first")

        try:
            document = Document(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        content_index = 0
        for paragraph in document.paragraphs:
            text = self._extract_paragraph_text(paragraph)
            if not text or text.strip() == "":
                continue
            content_index += 1
            segment_id = f"para_{content_index}"
            context = SegmentContext(
                file_path=filepath,
                location=f"Paragraph {content_index}",
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

    @staticmethod
    def _extract_paragraph_text(paragraph) -> str:
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        text_parts: list[str] = []
        for elem in paragraph._element.iter():
            if elem.tag == f"{{{word_ns}}}t":
                text_parts.append(elem.text or "")
            elif elem.tag == f"{{{word_ns}}}tab":
                text_parts.append("\t")
            elif elem.tag == f"{{{word_ns}}}br":
                text_parts.append("\n")
            elif elem.tag == f"{{{word_ns}}}cr":
                text_parts.append("\n")
        return "".join(text_parts)

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
