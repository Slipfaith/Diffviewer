from __future__ import annotations

from parsers.xliff_base import BaseXliffParser


class XliffParser(BaseXliffParser):
    name = "XLIFF Parser"
    supported_extensions = [".xliff", ".xlf"]
    format_description = "XLIFF"
