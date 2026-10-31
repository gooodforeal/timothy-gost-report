"""Loads and validates style.json: page and typography settings of the report.

Purpose:
    Defaults follow GOST. Whatever a methodical guide changes (margins, font sizes,
    caption numbering, first page number) goes into an optional style.json in the
    plain folder. Every key is validated, so typos fail loudly instead of silently
    falling back to defaults.

Components:
    ReportStyle: page and typography settings with GOST defaults.
    load_report_style: reads style.json, or returns defaults when it is absent.

Example:
    style = load_report_style(Path("plain/style.json"))
    print(style.font_size_pt)
"""

from __future__ import annotations

import json
from dataclasses import dataclass, fields
from pathlib import Path

CAPTION_NUMBERING_MODES = ("continuous", "by_section")


class StyleError(Exception):
    """Raised when style.json is malformed or has unknown keys or wrong types."""


@dataclass(frozen=True)
class ReportStyle:
    """Page and typography settings; field names are the keys of style.json."""

    font_name: str = "Times New Roman"
    font_size_pt: float = 14.0
    line_spacing: float = 1.5
    first_line_indent_cm: float = 1.25
    margin_left_mm: float = 30.0
    margin_right_mm: float = 15.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0
    listing_font_name: str = "Consolas"
    listing_font_size_pt: float = 12.0
    listing_line_spacing: float = 1.0
    table_font_size_pt: float = 12.0
    caption_numbering: str = "continuous"
    has_table_of_contents: bool = True
    # 2 when a separately made title page will be put in front of the report.
    first_page_number: int = 1


def load_report_style(style_path: Path) -> ReportStyle:
    """Reads style.json and merges it over the GOST defaults.

    Args:
        style_path: Path to style.json inside the plain folder; it may not exist.

    Returns:
        The validated style.

    Raises:
        StyleError: If the file is not valid JSON or has wrong keys or value types.
    """
    if not style_path.is_file():
        return ReportStyle()
    try:
        overrides = json.loads(style_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise StyleError(f"Invalid JSON in {style_path}: {error}") from error
    if not isinstance(overrides, dict):
        raise StyleError(f"{style_path.name} must be a JSON object")

    default_style = ReportStyle()
    default_by_key = {f.name: getattr(default_style, f.name) for f in fields(ReportStyle)}
    unknown_keys = sorted(set(overrides) - set(default_by_key))
    if unknown_keys:
        raise StyleError(f"Unknown keys in {style_path.name}: {', '.join(unknown_keys)}")

    # Each override must have the same type as its default; ints are accepted for floats.
    for key, value in overrides.items():
        default_value = default_by_key[key]
        is_number_for_float = isinstance(default_value, float) and type(value) in (int, float)
        if type(value) is not type(default_value) and not is_number_for_float:
            expected = type(default_value).__name__
            raise StyleError(f"{style_path.name}: {key} must be {expected}, got {value!r}")

    style = ReportStyle(**overrides)
    if style.caption_numbering not in CAPTION_NUMBERING_MODES:
        raise StyleError(f"{style_path.name}: caption_numbering must be one of {CAPTION_NUMBERING_MODES}")
    if style.first_page_number < 1:
        raise StyleError(f"{style_path.name}: first_page_number must be 1 or greater")
    return style
