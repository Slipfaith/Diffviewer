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


def test_memoq_strips_inline_tag_garbage(tmp_path: Path) -> None:
    mq_file = tmp_path / "with_ph.mqxliff"
    mq_file.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:mq="MQXliff">
  <file source-language="en" target-language="ru" datatype="plaintext" original="x.xlsx">
    <body>
      <trans-unit id="1" mq:status="NotStarted">
        <source xml:space="preserve"><ph id="1">&lt;mq:rxt displaytext="&amp;lt;p&amp;gt;" val="&amp;lt;p&amp;gt;"&gt;</ph>Hello<ph id="2">&lt;mq:rxt displaytext="&amp;lt;/p&amp;gt;" val="&amp;lt;/p&amp;gt;"&gt;</ph></source>
        <target xml:space="preserve"><ph id="1">&lt;mq:rxt displaytext="&amp;lt;p&amp;gt;" val="&amp;lt;p&amp;gt;"&gt;</ph>Привет<ph id="2">&lt;mq:rxt displaytext="&amp;lt;/p&amp;gt;" val="&amp;lt;/p&amp;gt;"&gt;</ph></target>
        <context-group>
          <context context-type="x-mmq-structural-context">Sheet:A1</context>
        </context-group>
        <mq:minorversions>
          <mq:historical-unit mq:status="NotStarted">
            <source xml:space="preserve">old source</source>
            <target xml:space="preserve">old target</target>
            <context-group>
              <context context-type="x-mmq-structural-context">OLD</context>
            </context-group>
            <note>old note</note>
          </mq:historical-unit>
        </mq:minorversions>
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
    seg = doc.segments[0]
    for field in (seg.source or "", seg.target):
        assert "mq:rxt" not in field
        assert "displaytext" not in field
        assert "&lt;" not in field
        assert "&amp;" not in field
    assert seg.source == "{1}Hello{2}"
    assert seg.target == "{1}Привет{2}"
    assert seg.metadata.get("context") == "Sheet:A1"
    assert seg.metadata.get("note") is None


def test_memoq_real_file_no_mq_noise() -> None:
    real = Path(__file__).resolve().parents[2] / "pptx" / "CM_Creator_Applications.xlsx_rus.mqxliff"
    if not real.exists():
        return
    parser = MemoQXliffParser()
    doc = parser.parse(str(real))
    assert doc.segments
    for seg in doc.segments:
        for field in (seg.source or "", seg.target):
            assert "mq:rxt" not in field, f"mq:rxt leaked into segment {seg.id}: {field!r}"
            assert "displaytext=" not in field, f"displaytext leaked into segment {seg.id}: {field!r}"
            assert "mq:ch" not in field, f"mq:ch leaked into segment {seg.id}: {field!r}"


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
