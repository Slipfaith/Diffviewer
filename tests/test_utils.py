from __future__ import annotations

from pathlib import Path

from core.utils import resource_path


def test_resource_path_dev_mode() -> None:
    template_rel = "reporters/templates/report.html.j2"
    absolute = Path(resource_path(template_rel))
    assert absolute.is_absolute()
    assert absolute.exists()
    assert absolute.name == "report.html.j2"

