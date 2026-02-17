from __future__ import annotations

from pathlib import Path

from core.utils import decode_html_entities, resource_path


def test_resource_path_dev_mode() -> None:
    template_rel = "reporters/templates/report.html.j2"
    absolute = Path(resource_path(template_rel))
    assert absolute.is_absolute()
    assert absolute.exists()
    assert absolute.name == "report.html.j2"


def test_decode_html_entities_decodes_single_and_nested() -> None:
    assert decode_html_entities("Don&#39;t") == "Don't"
    assert decode_html_entities("Can&amp;#39;t") == "Can't"


def test_decode_html_entities_can_preserve_single_encoded_literals() -> None:
    assert decode_html_entities("Don&#39;t", decode_single_encoded=False) == "Don&#39;t"
    assert decode_html_entities("Can&amp;#39;t", decode_single_encoded=False) == "Can't"
