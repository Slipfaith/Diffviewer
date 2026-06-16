from pathlib import Path

from core.models import ParseError
from core.registry import ParserRegistry


def _get_parser(path: Path):
    ParserRegistry._parsers.clear()
    ParserRegistry.discover()
    return ParserRegistry.get_parser(str(path))


def test_po_parse_header_multiline_entry_and_metadata(tmp_path: Path) -> None:
    path = tmp_path / "sample.po"
    path.write_text(
        '''msgid ""
msgstr ""
"Project-Id-Version: Sample\\n"
"Content-Type: text/plain; charset=utf-8\\n"

#. Extracted note
#: app.py:10
#, fuzzy, python-format
msgctxt "button"
msgid "Save"
msgstr "Save now"

msgid ""
"Line one "
"and two"
msgstr ""
"Target one\\n"
"Target two"
''',
        encoding="utf-8",
    )

    doc = _get_parser(path).parse(str(path))

    assert doc.format_name == "GNU gettext PO"
    assert doc.encoding == "utf-8-sig"
    assert doc.metadata["headers"]["Content-Type"] == "text/plain; charset=utf-8"
    assert len(doc.segments) == 2

    first = doc.segments[0]
    assert first.id.startswith("po:")
    assert first.source == "Save"
    assert first.target == "Save now"
    assert first.context.location == "entry 2"
    assert first.metadata["context"] == "button"
    assert first.metadata["references"] == ["app.py:10"]
    assert first.metadata["flags"] == ["fuzzy", "python-format"]
    assert first.metadata["extracted_comments"] == ["Extracted note"]

    second = doc.segments[1]
    assert second.source == "Line one and two"
    assert second.target == "Target one\nTarget two"


def test_po_parse_plural_msgstr_entries(tmp_path: Path) -> None:
    path = tmp_path / "plural.po"
    path.write_text(
        '''msgid "file"
msgid_plural "files"
msgstr[0] "one file"
msgstr[1] "many files"
''',
        encoding="utf-8",
    )

    doc = _get_parser(path).parse(str(path))

    assert [segment.source for segment in doc.segments] == ["file", "files"]
    assert [segment.target for segment in doc.segments] == ["one file", "many files"]
    assert [segment.metadata["plural_index"] for segment in doc.segments] == ["0", "1"]


def test_po_can_handle_extension_case_insensitively(tmp_path: Path) -> None:
    path = tmp_path / "sample.PO"
    path.write_text('msgid "Hello"\nmsgstr "Hi"\n', encoding="utf-8")

    parser = _get_parser(path)

    assert parser.can_handle(str(path)) is True
    assert parser.can_handle("sample.txt") is False


def test_po_invalid_quoted_string_raises_parse_error(tmp_path: Path) -> None:
    path = tmp_path / "bad.po"
    path.write_text('msgid Hello\nmsgstr "Hi"\n', encoding="utf-8")

    parser = _get_parser(path)

    assert parser.validate(str(path))
    try:
        parser.parse(str(path))
    except ParseError as exc:
        assert "quoted string" in exc.reason
    else:
        raise AssertionError("Expected ParseError")
