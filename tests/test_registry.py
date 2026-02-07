from __future__ import annotations

from pathlib import Path
import sys
import pytest

from core.models import ParsedDocument, UnsupportedFormatError
from core.registry import ParserRegistry, ReporterRegistry
from parsers.base import BaseParser
from reporters.base import BaseReporter


ROOT_DIR = Path(__file__).resolve().parents[1]
PARSERS_DIR = ROOT_DIR / "parsers"
REPORTERS_DIR = ROOT_DIR / "reporters"


def _write_module(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


@pytest.fixture()
def mock_plugins() -> None:
    ParserRegistry._parsers.clear()
    ReporterRegistry._reporters.clear()

    parser_module = PARSERS_DIR / "mock_parser_test.py"
    reporter_module = REPORTERS_DIR / "mock_reporter_test.py"

    _write_module(
        parser_module,
        """
from parsers.base import BaseParser
from core.models import ParsedDocument


class MockParser(BaseParser):
    name = "Mock Parser"
    supported_extensions = [".mock"]
    format_description = "Mock"

    def can_handle(self, filepath: str) -> bool:
        return filepath.lower().endswith(".mock")

    def parse(self, filepath: str) -> ParsedDocument:
        return ParsedDocument([], "MOCK", filepath)

    def validate(self, filepath: str) -> list[str]:
        return []
""".lstrip(),
    )

    _write_module(
        reporter_module,
        """
from reporters.base import BaseReporter
from core.models import ComparisonResult


class MockReporter(BaseReporter):
    name = "Mock Reporter"
    output_extension = ".mockout"
    supports_rich_text = False

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        return output_path + self.output_extension
""".lstrip(),
    )

    sys.modules.pop("parsers.mock_parser_test", None)
    sys.modules.pop("reporters.mock_reporter_test", None)

    yield

    if parser_module.exists():
        parser_module.unlink()
    if reporter_module.exists():
        reporter_module.unlink()

    sys.modules.pop("parsers.mock_parser_test", None)
    sys.modules.pop("reporters.mock_reporter_test", None)


def test_parser_discover_and_get_parser(mock_plugins: None) -> None:
    ParserRegistry.discover()
    assert ".mock" in ParserRegistry.supported_extensions()

    parser = ParserRegistry.get_parser("file.mock")
    assert parser.__class__.__name__ == "MockParser"

    with pytest.raises(UnsupportedFormatError):
        ParserRegistry.get_parser("file.unknown")


def test_parser_register_and_duplicates() -> None:
    ParserRegistry._parsers.clear()

    class RegisteredParser(BaseParser):
        name = "Registered"
        supported_extensions = [".reg"]
        format_description = "Registered"

        def can_handle(self, filepath: str) -> bool:
            return True

        def parse(self, filepath: str) -> ParsedDocument:
            return ParsedDocument([], "REG", filepath)

        def validate(self, filepath: str) -> list[str]:
            return []

    ParserRegistry.register(RegisteredParser)
    assert ".reg" in ParserRegistry.supported_extensions()

    class DuplicateParser(BaseParser):
        name = "Duplicate"
        supported_extensions = [".reg"]
        format_description = "Duplicate"

        def can_handle(self, filepath: str) -> bool:
            return True

        def parse(self, filepath: str) -> ParsedDocument:
            return ParsedDocument([], "REG", filepath)

        def validate(self, filepath: str) -> list[str]:
            return []

    with pytest.raises(ValueError):
        ParserRegistry.register(DuplicateParser)


def test_reporter_discover_and_get_reporter(mock_plugins: None) -> None:
    ReporterRegistry.discover()
    assert ".mockout" in ReporterRegistry.supported_extensions()

    reporter = ReporterRegistry.get_reporter(".mockout")
    assert reporter.__class__.__name__ == "MockReporter"

    with pytest.raises(UnsupportedFormatError):
        ReporterRegistry.get_reporter(".unknown")


def test_reporter_register_and_duplicates() -> None:
    ReporterRegistry._reporters.clear()

    class RegisteredReporter(BaseReporter):
        name = "Registered Reporter"
        output_extension = ".out"
        supports_rich_text = True

        def generate(self, result, output_path: str) -> str:
            return output_path + self.output_extension

    ReporterRegistry.register(RegisteredReporter)
    assert ".out" in ReporterRegistry.supported_extensions()

    class DuplicateReporter(BaseReporter):
        name = "Duplicate Reporter"
        output_extension = ".out"
        supports_rich_text = False

        def generate(self, result, output_path: str) -> str:
            return output_path + self.output_extension

    with pytest.raises(ValueError):
        ReporterRegistry.register(DuplicateReporter)
