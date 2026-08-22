"""Row/cell exclusion helpers used by lane parsers, built on excel_io.style_utils."""
from __future__ import annotations

from openpyxl.cell.cell import Cell

from mrg2opus.excel_io.style_utils import ExclusionConfig, is_excluded

__all__ = ["ExclusionConfig", "is_excluded", "location_is_excluded", "row_is_excluded"]


def row_is_excluded(cells: list[Cell], config: ExclusionConfig | None = None) -> bool:
    """True if every non-empty cell in the row is excluded (struck/blacked)."""
    non_empty = [c for c in cells if c.value not in (None, "")]
    if not non_empty:
        return False
    return all(is_excluded(c, config) for c in non_empty)


def location_is_excluded(cells: list[Cell], config: ExclusionConfig | None = None) -> bool:
    """True if ANY of a location's own name/code cells is struck-through or
    blacked-out - the raw-sheet convention for "this origin/destination is
    withdrawn, not included." Unlike row_is_excluded (which requires every
    cell excluded, appropriate for a row of many independent rate cells), a
    location is usually only 1-2 cells (name + code), and a trader marking
    just one of them - often just the code - is enough to mean the whole
    location is out; requiring both would silently keep withdrawn
    locations whenever the formatting wasn't applied perfectly to every
    cell."""
    non_empty = [c for c in cells if c.value not in (None, "")]
    if not non_empty:
        return False
    return any(is_excluded(c, config) for c in non_empty)
