from __future__ import annotations

from pathlib import Path

import openpyxl

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


class XlsxParser(BaseParser):
    name = "XLSX Parser"
    supported_extensions = [".xlsx"]
    format_description = "Excel Workbook"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            workbook = openpyxl.load_workbook(filepath, data_only=True)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if value is None or value == "":
                        continue
                    cell_id = f"{sheet.title}!{cell.coordinate}"
                    context = SegmentContext(
                        file_path=filepath,
                        location=cell_id,
                        position=len(segments) + 1,
                        group=sheet.title,
                    )
                    segments.append(
                        Segment(
                            id=cell_id,
                            source=None,
                            target=str(value),
                            context=context,
                            metadata={
                                "row": cell.row,
                                "column": cell.column,
                                "sheet_name": sheet.title,
                            },
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
            _ = openpyxl.load_workbook(filepath, data_only=True)
        except Exception as exc:
            errors.append(str(exc))
        return errors
