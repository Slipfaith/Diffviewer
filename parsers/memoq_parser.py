from __future__ import annotations

from lxml import etree

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from core.utils import decode_html_entities
from parsers.xliff_base import BaseXliffParser, _first_child_text, _local_name


class MemoQXliffParser(BaseXliffParser):
    name = "MemoQ XLIFF Parser"
    supported_extensions = [".mqxliff"]
    format_description = "MemoQ XLIFF"

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            parser = etree.XMLParser(resolve_entities=False, recover=False)
            tree = etree.parse(filepath, parser)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        root = tree.getroot()
        segments: list[Segment] = []

        trans_units = root.xpath(".//*[local-name()='trans-unit']")
        for index, unit in enumerate(trans_units, start=1):
            unit_id = unit.get("id") or str(index)
            source = _first_child_text(unit, "source")
            target = _first_child_text(unit, "target") or ""

            metadata: dict[str, str] = {}
            for attr_name, attr_value in unit.attrib.items():
                name = _local_name(attr_name)
                if name in {
                    "status",
                    "segmentguid",
                    "lastchanginguser",
                    "lastchangedtimestamp",
                }:
                    metadata[name] = attr_value

            context_nodes = unit.xpath(
                ".//*[local-name()='context' and @context-type='x-mmq-structural-context']"
            )
            if context_nodes:
                metadata["context"] = decode_html_entities(
                    "".join(context_nodes[0].itertext()).strip()
                )

            note_nodes = unit.xpath(".//*[local-name()='note']")
            if note_nodes:
                metadata["note"] = decode_html_entities(
                    "".join(note_nodes[0].itertext()).strip()
                )

            context = SegmentContext(
                file_path=filepath,
                location=unit_id,
                position=len(segments) + 1,
                group=None,
            )
            segments.append(
                Segment(
                    id=unit_id,
                    source=source,
                    target=target,
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
