from __future__ import annotations

from pathlib import Path

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class PdfParser(BaseParser):
    name = "PDF Parser"
    supported_extensions = [".pdf"]
    format_description = "PDF Document"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            import fitz  # PyMuPDF
        except ImportError as exc:
            raise ParseError(filepath, "PyMuPDF (fitz) is not installed. Run: pip install pymupdf") from exc

        try:
            doc = fitz.open(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        try:
            for page_num, page in enumerate(doc, start=1):
                blocks = page.get_text("blocks")
                para_num = 0
                for block in blocks:
                    # block: (x0, y0, x1, y1, text, block_no, block_type)
                    # block_type 0 = text, 1 = image
                    if block[6] != 0:
                        continue
                    text = block[4].strip()
                    if not text:
                        continue
                    para_num += 1
                    segment_id = f"page{page_num}_para{para_num}"
                    context = SegmentContext(
                        file_path=filepath,
                        location=f"Page {page_num}, paragraph {para_num}",
                        position=para_num,
                        group=f"Page {page_num}",
                    )
                    segments.append(
                        Segment(
                            id=segment_id,
                            source=None,
                            target=text,
                            context=context,
                        )
                    )
        finally:
            doc.close()

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={},
            encoding="utf-8",
        )

    def validate(self, filepath: str) -> list[str]:
        errors: list[str] = []
        try:
            import fitz
        except ImportError:
            errors.append("PyMuPDF (fitz) is not installed. Run: pip install pymupdf")
            return errors
        try:
            doc = fitz.open(filepath)
            doc.close()
        except Exception as exc:
            errors.append(str(exc))
        return errors
