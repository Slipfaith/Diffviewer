from __future__ import annotations

import importlib
import inspect
from pathlib import Path
from typing import Callable, Type

from core.models import UnsupportedFormatError
from parsers.base import BaseParser
from reporters.base import BaseReporter


PARSER_HIDDEN_IMPORTS = [
    "parsers.xliff_parser",
    "parsers.sdlxliff_parser",
    "parsers.memoq_parser",
    "parsers.txt_parser",
    "parsers.srt_parser",
    "parsers.xlsx_parser",
    "parsers.xls_parser",
    "parsers.pptx_parser",
    "parsers.docx_parser",
]

REPORTER_HIDDEN_IMPORTS = [
    "reporters.html_reporter",
    "reporters.excel_reporter",
    "reporters.docx_reporter",
    "reporters.summary_reporter",
]


class ParserRegistry:
    _parsers: dict[str, Type[BaseParser]] = {}

    @classmethod
    def discover(cls) -> None:
        cls._discover_package("parsers", BaseParser, cls.register)
        if not cls._parsers:
            cls._register_known_modules(PARSER_HIDDEN_IMPORTS, BaseParser, cls.register)

    @classmethod
    def register(cls, parser_class: Type[BaseParser]) -> None:
        if not inspect.isclass(parser_class) or not issubclass(parser_class, BaseParser):
            raise TypeError("Parser must be a BaseParser subclass")
        for ext in parser_class.supported_extensions:
            key = ext.lower()
            existing = cls._parsers.get(key)
            if existing is not None and existing is not parser_class:
                raise ValueError(f"Extension already registered: {key}")
            cls._parsers[key] = parser_class

    @classmethod
    def get_parser(cls, filepath: str) -> BaseParser:
        ext = Path(filepath).suffix.lower()
        parser_class = cls._parsers.get(ext)
        if parser_class is None:
            raise UnsupportedFormatError(ext)
        parser = parser_class()
        if not parser.can_handle(filepath):
            raise UnsupportedFormatError(ext)
        return parser

    @classmethod
    def supported_extensions(cls) -> list[str]:
        return sorted(cls._parsers.keys())

    @staticmethod
    def _discover_package(
        package_name: str,
        base_class: Type[BaseParser],
        register: Callable,
    ) -> None:
        package_dir = Path(__file__).resolve().parents[1] / package_name
        if not package_dir.exists():
            return
        for path in package_dir.glob("*.py"):
            if path.name.startswith("_"):
                continue
            module_name = f"{package_name}.{path.stem}"
            module = importlib.import_module(module_name)
            ParserRegistry._register_module_classes(module, base_class, register)

    @staticmethod
    def _register_known_modules(
        module_names: list[str],
        base_class: Type[BaseParser],
        register: Callable,
    ) -> None:
        for module_name in module_names:
            module = importlib.import_module(module_name)
            ParserRegistry._register_module_classes(module, base_class, register)

    @staticmethod
    def _register_module_classes(module, base_class: Type[BaseParser], register: Callable) -> None:
        for _, obj in inspect.getmembers(module, inspect.isclass):
            if obj is base_class or not issubclass(obj, base_class):
                continue
            if inspect.isabstract(obj):
                continue
            register(obj)


class ReporterRegistry:
    _reporters: dict[str, Type[BaseReporter]] = {}

    @classmethod
    def discover(cls) -> None:
        cls._discover_package("reporters", BaseReporter, cls.register)
        if not cls._reporters:
            cls._register_known_modules(
                REPORTER_HIDDEN_IMPORTS,
                BaseReporter,
                cls.register,
            )

    @classmethod
    def register(cls, reporter_class: Type[BaseReporter]) -> None:
        if not inspect.isclass(reporter_class) or not issubclass(reporter_class, BaseReporter):
            raise TypeError("Reporter must be a BaseReporter subclass")
        key = reporter_class.output_extension.lower()
        existing = cls._reporters.get(key)
        if existing is not None and existing is not reporter_class:
            raise ValueError(f"Extension already registered: {key}")
        cls._reporters[key] = reporter_class

    @classmethod
    def get_reporter(cls, output_extension: str) -> BaseReporter:
        key = output_extension.lower()
        reporter_class = cls._reporters.get(key)
        if reporter_class is None:
            raise UnsupportedFormatError(key)
        return reporter_class()

    @classmethod
    def supported_extensions(cls) -> list[str]:
        return sorted(cls._reporters.keys())

    @staticmethod
    def _discover_package(
        package_name: str,
        base_class: Type[BaseReporter],
        register: Callable,
    ) -> None:
        package_dir = Path(__file__).resolve().parents[1] / package_name
        if not package_dir.exists():
            return
        for path in package_dir.glob("*.py"):
            if path.name.startswith("_"):
                continue
            module_name = f"{package_name}.{path.stem}"
            module = importlib.import_module(module_name)
            ReporterRegistry._register_module_classes(module, base_class, register)

    @staticmethod
    def _register_known_modules(
        module_names: list[str],
        base_class: Type[BaseReporter],
        register: Callable,
    ) -> None:
        for module_name in module_names:
            module = importlib.import_module(module_name)
            ReporterRegistry._register_module_classes(module, base_class, register)

    @staticmethod
    def _register_module_classes(module, base_class: Type[BaseReporter], register: Callable) -> None:
        for _, obj in inspect.getmembers(module, inspect.isclass):
            if obj is base_class or not issubclass(obj, base_class):
                continue
            if inspect.isabstract(obj):
                continue
            register(obj)

