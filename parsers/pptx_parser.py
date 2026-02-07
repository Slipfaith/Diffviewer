from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class PptxParser(BaseParser):
    name = "PPTX Parser"
    supported_extensions = [".pptx"]
    format_description = "PowerPoint Presentation"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            presentation = Presentation(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        for slide_index, slide in enumerate(presentation.slides, start=1):
            for shape_index, shape in enumerate(slide.shapes, start=1):
                if not shape.has_text_frame:
                    continue
                for para_index, paragraph in enumerate(shape.text_frame.paragraphs, start=1):
                    text = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
                    if not text or text.strip() == "":
                        continue
                    segment_id = f"slide{slide_index}_shape{shape_index}_para{para_index}"
                    context = SegmentContext(
                        file_path=filepath,
                        location=f"Slide {slide_index} > {shape.name}",
                        position=len(segments) + 1,
                        group=f"Slide {slide_index}",
                    )
                    segments.append(
                        Segment(
                            id=segment_id,
                            source=None,
                            target=text,
                            context=context,
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
            _ = Presentation(filepath)
        except Exception as exc:
            errors.append(str(exc))
        return errors
