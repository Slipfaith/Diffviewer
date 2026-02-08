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

        def add_segment(segment_id: str, location: str, group: str, text: str) -> None:
            normalized = (text or "").strip()
            if not normalized:
                return
            context = SegmentContext(
                file_path=filepath,
                location=location,
                position=len(segments) + 1,
                group=group,
            )
            segments.append(
                Segment(
                    id=segment_id,
                    source=None,
                    target=normalized,
                    context=context,
                )
            )

        def extract_shape_text(shape, *, slide_index: int, shape_index: int, notes: bool = False) -> None:
            area_name = "Notes" if notes else "Slide"
            group = f"Slide {slide_index} Notes" if notes else f"Slide {slide_index}"
            shape_name = getattr(shape, "name", "Shape")

            if shape.has_text_frame:
                for para_index, paragraph in enumerate(shape.text_frame.paragraphs, start=1):
                    text = (
                        "".join(run.text for run in paragraph.runs)
                        if paragraph.runs
                        else paragraph.text
                    )
                    segment_id = (
                        f"slide{slide_index}_notes_shape{shape_index}_para{para_index}"
                        if notes
                        else f"slide{slide_index}_shape{shape_index}_para{para_index}"
                    )
                    location = f"Slide {slide_index} > {area_name} > {shape_name}"
                    add_segment(segment_id, location, group, text)

            if shape.has_table:
                for row_index, row in enumerate(shape.table.rows, start=1):
                    for cell_index, cell in enumerate(row.cells, start=1):
                        segment_id = (
                            f"slide{slide_index}_notes_shape{shape_index}_tbl_r{row_index}c{cell_index}"
                            if notes
                            else f"slide{slide_index}_shape{shape_index}_tbl_r{row_index}c{cell_index}"
                        )
                        location = (
                            f"Slide {slide_index} > {area_name} > {shape_name} > "
                            f"Table R{row_index}C{cell_index}"
                        )
                        add_segment(segment_id, location, group, cell.text)

        for slide_index, slide in enumerate(presentation.slides, start=1):
            for shape_index, shape in enumerate(slide.shapes, start=1):
                extract_shape_text(
                    shape,
                    slide_index=slide_index,
                    shape_index=shape_index,
                )

            if getattr(slide, "has_notes_slide", False):
                try:
                    notes_slide = slide.notes_slide
                except Exception:
                    notes_slide = None
                if notes_slide is not None:
                    for shape_index, shape in enumerate(notes_slide.shapes, start=1):
                        extract_shape_text(
                            shape,
                            slide_index=slide_index,
                            shape_index=shape_index,
                            notes=True,
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
