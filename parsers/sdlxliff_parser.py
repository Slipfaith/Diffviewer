from __future__ import annotations

from pathlib import Path

from lxml import etree

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from core.utils import decode_html_entities
from parsers.xliff_base import BaseXliffParser


def _iter_mrk_segments(element: etree._Element) -> list[etree._Element]:
    return element.xpath(".//*[local-name()='mrk' and @mtype='seg']")


def _extract_text(node: etree._Element) -> str:
    return decode_html_entities("".join(node.itertext()))


class SdlXliffParser(BaseXliffParser):
    name = "SDLXLIFF Parser"
    supported_extensions = [".sdlxliff"]
    format_description = "SDLXLIFF"

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            parser = etree.XMLParser(resolve_entities=False, recover=False)
            tree = etree.parse(filepath, parser)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        root = tree.getroot()
        segments: list[Segment] = []

        trans_units = root.xpath(".//*[local-name()='trans-unit']")
        for unit in trans_units:
            seg_source = unit.xpath(".//*[local-name()='seg-source']")
            target = unit.xpath(".//*[local-name()='target']")
            seg_source_node = seg_source[0] if seg_source else None
            target_node = target[0] if target else None

            if seg_source_node is None:
                continue

            source_mrks = _iter_mrk_segments(seg_source_node)
            target_mrks = _iter_mrk_segments(target_node) if target_node is not None else []
            target_map = {mrk.get("mid"): mrk for mrk in target_mrks if mrk.get("mid")}

            seg_defs = unit.xpath(".//*[local-name()='seg-defs']//*[local-name()='seg']")
            seg_meta = {}
            for seg in seg_defs:
                seg_id = seg.get("id")
                if not seg_id:
                    continue
                seg_meta[seg_id] = {k: v for k, v in seg.attrib.items() if k != "id"}

            for index, mrk in enumerate(source_mrks, start=1):
                mid = mrk.get("mid") or str(index)
                source_text = _extract_text(mrk)
                target_mrk = target_map.get(mid)
                target_text = _extract_text(target_mrk) if target_mrk is not None else ""
                context = SegmentContext(
                    file_path=filepath,
                    location=mid,
                    position=len(segments) + 1,
                    group=None,
                )
                metadata = seg_meta.get(mid, {})
                segments.append(
                    Segment(
                        id=mid,
                        source=source_text,
                        target=target_text,
                        context=context,
                        metadata=dict(metadata),
                    )
                )

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={},
            encoding=None,
        )
