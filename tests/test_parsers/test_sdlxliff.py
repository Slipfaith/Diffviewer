from pathlib import Path

from parsers.sdlxliff_parser import SdlXliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_sdlxliff_parse_fixture() -> None:
    parser = SdlXliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.sdlxliff"))
    assert len(doc.segments) == 3
    assert doc.segments[0].id == "101"
    assert doc.segments[0].source == "First line"
    assert doc.segments[0].target == "Первая строка"
    assert doc.segments[0].metadata.get("conf") == "ApprovedTranslation"


def test_sdlxliff_can_handle() -> None:
    parser = SdlXliffParser()
    assert parser.can_handle("file.sdlxliff") is True
    assert parser.can_handle("file.xliff") is False


def test_sdlxliff_decodes_nested_html_entities(tmp_path: Path) -> None:
    sdl_file = tmp_path / "entities.sdlxliff"
    sdl_file.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:sdl="http://sdl.com/FileTypes/SdlXliff/1.0">
  <file source-language="en" target-language="en" datatype="plaintext" original="sample.txt">
    <body>
      <trans-unit id="tu-1">
        <seg-source><g id="1"><mrk mtype="seg" mid="101">Don&amp;#39;t</mrk></g></seg-source>
        <target><g id="1"><mrk mtype="seg" mid="101">Can&amp;#39;t</mrk></g></target>
        <sdl:seg-defs>
          <sdl:seg id="101" conf="ApprovedTranslation"/>
        </sdl:seg-defs>
      </trans-unit>
    </body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )

    parser = SdlXliffParser()
    doc = parser.parse(str(sdl_file))

    assert len(doc.segments) == 1
    assert doc.segments[0].source == "Don't"
    assert doc.segments[0].target == "Can't"
