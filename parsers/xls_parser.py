from __future__ import annotations

from pathlib import Path

import xlrd

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


def _col_to_letters(col_index: int) -> str:
    letters = ""
    index = col_index + 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


class XlsParser(BaseParser):
    name = "XLS Parser"
    supported_extensions = [".xls"]
    format_description = "Excel Workbook (XLS)"

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            workbook = xlrd.open_workbook(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        for sheet in workbook.sheets():
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    value = sheet.cell_value(row, col)
                    if value == "" or value is None:
                        continue
                    coordinate = f"{_col_to_letters(col)}{row + 1}"
                    cell_id = f"{sheet.name}!{coordinate}"
                    context = SegmentContext(
                        file_path=filepath,
                        location=cell_id,
                        position=len(segments) + 1,
                        group=sheet.name,
                    )
                    segments.append(
                        Segment(
                            id=cell_id,
                            source=None,
                            target=str(value),
                            context=context,
                            metadata={
                                "row": row + 1,
                                "column": col + 1,
                                "sheet_name": sheet.name,
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
            _ = xlrd.open_workbook(filepath)
        except Exception as exc:
            errors.append(str(exc))
        return errors
