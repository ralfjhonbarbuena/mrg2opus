"""Style-aware cell inspection: detect cells that must be excluded from
extraction because the raw MRG marks them as cancelled/superseded (strikethrough)
or blacked-out (fill color), per the source MRG convention.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from functools import lru_cache

from openpyxl.cell.cell import Cell

# Colors traders commonly use to "black out" a cancelled cell. Kept
# configurable (not a single hardcoded RGB) since this varies by trader.
DEFAULT_BLACKOUT_RGBS = frozenset({
    "FF000000",  # black
    "00000000",  # black, no alpha
    "FF808080",  # grey
    "FFFF0000",  # red (also used for struck-through text color in samples, but
                 # a red *fill* specifically usually means "cancelled")
})

# A cell's fill can reference a WORKBOOK THEME color swatch (Excel's Fill
# Color -> theme palette square) instead of a literal RGB - confirmed
# real-world case: a lane's raw sheet zeroed out one origin's rates across
# every destination by blacking the row out via exactly this styling, not
# a literal black RGB fill, and silently filed real $0.00 rates for all of
# them until this was recognized. OOXML SpreadsheetML's theme color INDEX
# (as referenced by a cell fill) is swapped relative to the theme XML's
# own <a:clrScheme> declaration order for the first two background/text
# pairs - a documented, consistent OOXML quirk, verified against the real
# file that triggered this (its theme declares dk1/dk2=black, lt1/lt2=
# white; a cell with theme index 1 there renders visibly black, matching
# dk1 - not lt1, which the unswapped declaration order would give).
THEME_COLOR_ORDER = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]
_THEME_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


@dataclass(frozen=True)
class ExclusionConfig:
    blackout_rgbs: frozenset[str] = field(default_factory=lambda: DEFAULT_BLACKOUT_RGBS)
    treat_strikethrough_as_excluded: bool = True
    treat_fill_as_excluded: bool = True


def is_struck_through(cell: Cell) -> bool:
    font = cell.font
    return bool(font is not None and font.strike)


@lru_cache(maxsize=8)
def _resolve_theme_rgb(theme_xml: bytes, theme_index: int) -> str | None:
    """Resolve a fill's `theme` color index to its 8-char ARGB hex (e.g.
    "FF000000"), by parsing the WORKBOOK's own theme XML - themes are
    frequently customized per file, never assume a fixed palette. `tint`
    (a lighten/darken adjustment some themed fills also carry) is not
    applied - every real case seen so far uses tint=0 (the theme color
    exactly), and precise tint math is unverified guesswork without a
    real example that needs it. Cached per (theme bytes, index) pair since
    the same workbook's theme is parsed on every excluded cell otherwise."""
    if theme_index < 0 or theme_index >= len(THEME_COLOR_ORDER):
        return None
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError:
        return None
    color_scheme = root.find(".//a:clrScheme", _THEME_NS)
    if color_scheme is None:
        return None
    element = color_scheme.find(f"a:{THEME_COLOR_ORDER[theme_index]}", _THEME_NS)
    if element is None:
        return None
    srgb = element.find("a:srgbClr", _THEME_NS)
    if srgb is not None and srgb.get("val"):
        return f"FF{srgb.get('val').upper()}"
    sys_clr = element.find("a:sysClr", _THEME_NS)
    if sys_clr is not None and sys_clr.get("lastClr"):
        return f"FF{sys_clr.get('lastClr').upper()}"
    return None


def is_blacked_out(cell: Cell, config: ExclusionConfig | None = None) -> bool:
    config = config or ExclusionConfig()
    fill = cell.fill
    if fill is None or fill.fill_type is None:
        return False
    color = fill.start_color
    if color is None:
        return False
    if color.type == "rgb":
        return color.rgb in config.blackout_rgbs
    if color.type == "theme":
        worksheet = getattr(cell, "parent", None)
        workbook = getattr(worksheet, "parent", None)
        theme_xml = getattr(workbook, "loaded_theme", None)
        if not theme_xml:
            return False
        resolved = _resolve_theme_rgb(theme_xml, color.theme)
        return resolved is not None and resolved in config.blackout_rgbs
    return False


def is_excluded(cell: Cell, config: ExclusionConfig | None = None) -> bool:
    """True if this cell's value should be dropped from extraction."""
    config = config or ExclusionConfig()
    if config.treat_strikethrough_as_excluded and is_struck_through(cell):
        return True
    if config.treat_fill_as_excluded and is_blacked_out(cell, config):
        return True
    return False
