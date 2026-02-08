from __future__ import annotations

from pathlib import Path

import xlrd

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


def _parse_column_reference(column: str | int | None) -> int | None:
    if column is None:
        return None
    if isinstance(column, int):
        if column < 1:
            raise ValueError("Excel source column must be >= 1.")
        return column

    raw = str(column).strip()
    if not raw:
        return None
    if raw.isdigit():
        value = int(raw)
        if value < 1:
            raise ValueError("Excel source column must be >= 1.")
        return value
    if raw.isalpha():
        value = 0
        for char in raw.upper():
            value = value * 26 + (ord(char) - 64)
        return value
    raise ValueError(f"Invalid Excel source column: {column}")


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

    def __init__(self) -> None:
        self._source_column_index: int | None = None

    def can_handle(self, filepath: str) -> bool:
        return Path(filepath).suffix.lower() in self.supported_extensions

    def set_source_column(self, column: str | int | None) -> None:
        self._source_column_index = _parse_column_reference(column)

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            workbook = xlrd.open_workbook(filepath)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        segments: list[Segment] = []
        source_column_index = self._source_column_index
        for sheet in workbook.sheets():
            for row in range(sheet.nrows):
                source_value = self._read_row_source_value(sheet, row, source_column_index)
                for col in range(sheet.ncols):
                    value = sheet.cell_value(row, col)
                    if value == "" or value is None:
                        continue
                    if source_column_index is not None and col == source_column_index - 1:
                        continue
                    coordinate = f"{_col_to_letters(col)}{row + 1}"
                    cell_id = f"{sheet.name}!{coordinate}"
                    context = SegmentContext(
                        file_path=filepath,
                        location=cell_id,
                        position=len(segments) + 1,
                        group=sheet.name,
                    )
                    metadata = {
                        "row": row + 1,
                        "column": col + 1,
                        "sheet_name": sheet.name,
                    }
                    if source_column_index is not None:
                        metadata["source_column"] = source_column_index
                    segments.append(
                        Segment(
                            id=cell_id,
                            source=source_value,
                            target=str(value),
                            context=context,
                            metadata=metadata,
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

    @staticmethod
    def _read_row_source_value(
        sheet,
        row: int,
        source_column_index: int | None,
    ) -> str | None:
        if source_column_index is None:
            return None
        source_column_zero_based = source_column_index - 1
        if source_column_zero_based < 0 or source_column_zero_based >= sheet.ncols:
            return None
        value = sheet.cell_value(row, source_column_zero_based)
        if value is None or value == "":
            return None
        text = str(value)
        return text if text != "" else None
