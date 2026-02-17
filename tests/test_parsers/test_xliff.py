from pathlib import Path

import pytest

from core.models import ParseError
from parsers.xliff_parser import XliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_xliff_parse_fixture_a() -> None:
    parser = XliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.xliff"))
    assert len(doc.segments) == 4
    assert [seg.id for seg in doc.segments] == ["1", "2", "3", "4"]
    assert doc.segments[0].source == "Hello"
    assert doc.segments[0].target == "Привет"


def test_xliff_parse_fixture_b() -> None:
    parser = XliffParser()
    doc = parser.parse(str(FIXTURES / "sample_b.xliff"))
    assert len(doc.segments) == 4
    ids = [seg.id for seg in doc.segments]
    assert "5" in ids
    target_map = {seg.id: seg.target for seg in doc.segments}
    assert target_map["2"] == "Мир!"


def test_xliff_can_handle() -> None:
    parser = XliffParser()
    assert parser.can_handle("file.xliff") is True
    assert parser.can_handle("file.xlf") is True
    assert parser.can_handle("file.txt") is False


def test_xliff_empty_file(tmp_path: Path) -> None:
    empty_file = tmp_path / "empty.xliff"
    empty_file.write_text("", encoding="utf-8")
    parser = XliffParser()
    with pytest.raises(ParseError):
        parser.parse(str(empty_file))


def test_xliff_decodes_nested_html_entities(tmp_path: Path) -> None:
    xliff_file = tmp_path / "entities.xliff"
    xliff_file.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file source-language="en" target-language="en" datatype="plaintext" original="sample.txt">
    <body>
      <trans-unit id="1">
        <source>Don&amp;#39;t</source>
        <target>Can&amp;#39;t</target>
      </trans-unit>
    </body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )

    parser = XliffParser()
    doc = parser.parse(str(xliff_file))

    assert len(doc.segments) == 1
    assert doc.segments[0].source == "Don't"
    assert doc.segments[0].target == "Can't"
