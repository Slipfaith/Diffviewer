from __future__ import annotations

import html
import os
import sys


def resource_path(relative_path: str) -> str:
    """Return absolute path to resource for dev and PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)  # type: ignore[attr-defined]
    return os.path.join(os.path.abspath("."), relative_path)


def decode_html_entities(
    value: str,
    *,
    max_rounds: int = 4,
    decode_single_encoded: bool = True,
) -> str:
    """Decode HTML entities, optionally preserving single-encoded literals."""
    if not decode_single_encoded and "&amp;" not in value:
        return value

    current = value
    for _ in range(max_rounds):
        decoded = html.unescape(current)
        if decoded == current:
            break
        current = decoded
    return current
