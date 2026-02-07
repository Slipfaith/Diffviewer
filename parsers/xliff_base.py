from __future__ import annotations

from pathlib import Path

from lxml import etree

from core.models import ParsedDocument, ParseError, Segment, SegmentContext
from parsers.base import BaseParser


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _first_child_text(element: etree._Element, name: str) -> str | None:
    matches = element.xpath(f".//*[local-name()='{name}']")
    if not matches:
        return None
    text = "".join(matches[0].itertext())
    return text


class BaseXliffParser(BaseParser):
    name = "Base XLIFF Parser"
    supported_extensions: list[str] = []
    format_description = "XLIFF"

    def can_handle(self, filepath: str) -> bool:
        ext = Path(filepath).suffix.lower()
        if ext not in self.supported_extensions:
            return False
        path = Path(filepath)
        if not path.exists():
            return True
        try:
            for _, elem in etree.iterparse(str(path), events=("start",), recover=True):
                return _local_name(elem.tag) == "xliff"
        except Exception:
            return False
        return False

    def parse(self, filepath: str) -> ParsedDocument:
        try:
            parser = etree.XMLParser(resolve_entities=False, recover=False)
            tree = etree.parse(filepath, parser)
        except Exception as exc:
            raise ParseError(filepath, str(exc)) from exc

        root = tree.getroot()
        segments: list[Segment] = []

        trans_units = root.xpath(".//*[local-name()='trans-unit']")
        if trans_units:
            for index, unit in enumerate(trans_units, start=1):
                unit_id = unit.get("id") or str(index)
                source = _first_child_text(unit, "source")
                target = _first_child_text(unit, "target") or ""
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
                    )
                )
        else:
            units = root.xpath(".//*[local-name()='unit']")
            for unit in units:
                unit_id = unit.get("id")
                segments_in_unit = unit.xpath(".//*[local-name()='segment']")
                if segments_in_unit:
                    for index, segment in enumerate(segments_in_unit, start=1):
                        seg_id = segment.get("id") or unit_id or str(index)
                        if unit_id and segment.get("id"):
                            final_id = f"{unit_id}:{seg_id}"
                        else:
                            final_id = seg_id
                        source = _first_child_text(segment, "source")
                        target = _first_child_text(segment, "target") or ""
                        context = SegmentContext(
                            file_path=filepath,
                            location=final_id,
                            position=len(segments) + 1,
                            group=None,
                        )
                        segments.append(
                            Segment(
                                id=final_id,
                                source=source,
                                target=target,
                                context=context,
                            )
                        )
                elif unit_id is not None:
                    source = _first_child_text(unit, "source")
                    target = _first_child_text(unit, "target") or ""
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
                        )
                    )

        return ParsedDocument(
            segments=segments,
            format_name=self.format_description,
            file_path=filepath,
            metadata={},
            encoding=None,
        )

    def validate(self, filepath: str) -> list[str]:
        errors: list[str] = []
        try:
            parser = etree.XMLParser(resolve_entities=False, recover=False)
            tree = etree.parse(filepath, parser)
            root = tree.getroot()
            if _local_name(root.tag) != "xliff":
                errors.append("Root element is not xliff")
        except Exception as exc:
            errors.append(str(exc))
        return errors
