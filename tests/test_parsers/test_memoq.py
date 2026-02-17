from pathlib import Path

from parsers.memoq_parser import MemoQXliffParser


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_memoq_parse_fixture() -> None:
    parser = MemoQXliffParser()
    doc = parser.parse(str(FIXTURES / "sample_a.mqxliff"))
    assert len(doc.segments) == 3
    seg = doc.segments[0]
    assert seg.id == "1"
    assert seg.source.startswith("Log in")
    assert seg.target.startswith("Log elke")
    assert seg.metadata.get("status") == "ManuallyConfirmed"
    assert seg.metadata.get("segmentguid") == "guid-1"
    assert seg.metadata.get("context") == ":B5/C5"


def test_memoq_can_handle() -> None:
    parser = MemoQXliffParser()
    assert parser.can_handle("file.mqxliff") is True
    assert parser.can_handle("file.xliff") is False


def test_memoq_decodes_nested_html_entities(tmp_path: Path) -> None:
    mq_file = tmp_path / "entities.mqxliff"
    mq_file.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:mq="http://memoq.com/xliff">
  <file source-language="en" target-language="en" datatype="plaintext" original="sample.txt">
    <body>
      <trans-unit id="1" mq:status="Translated">
        <source>Don&amp;#39;t</source>
        <target>Can&amp;#39;t</target>
        <context-group>
          <context context-type="x-mmq-structural-context">A&amp;#39;B</context>
        </context-group>
        <note>N&amp;#39;1</note>
      </trans-unit>
    </body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )

    parser = MemoQXliffParser()
    doc = parser.parse(str(mq_file))

    assert len(doc.segments) == 1
    assert doc.segments[0].source == "Don't"
    assert doc.segments[0].target == "Can't"
    assert doc.segments[0].metadata.get("context") == "A'B"
    assert doc.segments[0].metadata.get("note") == "N'1"
