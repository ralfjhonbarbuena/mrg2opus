"""Style-aware cell inspection: detect cells that must be excluded from
extraction because the raw MRG marks them as cancelled/superseded (strikethrough)
or blacked-out (fill color), per the source MRG convention.
"""
from __future__ import annotations

from dataclasses import dataclass, field
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


@dataclass(frozen=True)
class ExclusionConfig:
    blackout_rgbs: frozenset[str] = field(default_factory=lambda: DEFAULT_BLACKOUT_RGBS)
    treat_strikethrough_as_excluded: bool = True
    treat_fill_as_excluded: bool = True


def is_struck_through(cell: Cell) -> bool:
    font = cell.font
    return bool(font is not None and font.strike)


def is_blacked_out(cell: Cell, config: ExclusionConfig | None = None) -> bool:
    config = config or ExclusionConfig()
    fill = cell.fill
    if fill is None or fill.fill_type is None:
        return False
    color = fill.start_color
    if color is None or color.type != "rgb":
        return False
    return color.rgb in config.blackout_rgbs


def is_excluded(cell: Cell, config: ExclusionConfig | None = None) -> bool:
    """True if this cell's value should be dropped from extraction."""
    config = config or ExclusionConfig()
    if config.treat_strikethrough_as_excluded and is_struck_through(cell):
        return True
    if config.treat_fill_as_excluded and is_blacked_out(cell, config):
        return True
    return False
