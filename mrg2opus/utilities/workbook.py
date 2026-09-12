"""Reading an OPUS workbook back in, whatever its sheets are called.

The 45 real filings in reference/ spell one sheet five different ways -
the surcharge sheet appears as SRCHG (19 files), CMDT NOTE (18), SUR (9),
SURCHARGE (2) and sur (1), and the route-note sheet as RN, ROUTE NOTE,
ROUTE, RNT and rnt. find_sheet() already ignores case, whitespace and
hyphens; the aliases below are the rest of it, so a filing opens whoever
prepared it.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from openpyxl.workbook import Workbook

from mrg2opus.audit.compare import (
    find_sheet,
    read_arbs_sheet,
    read_cmdt_note_sheet,
    read_freetime_sheet,
    read_rates_sheet,
    read_route_note_sheet,
    read_special_note_sheet,
    read_vertical_rates_sheet,
)

# Every name a real filing has used for each sheet, canonical first. The
# counts above come from reference/2_OPUS; DEMDET and O.ARBS are the
# names recorded in the note-sheet taxonomy but not yet seen in a file.
SHEET_ALIASES: dict[str, tuple[str, ...]] = {
    "rates": ("RATES",),
    "rates_port_port": ("RATES PORT-PORT", "PORT-PORT", "RATES PORT - PORT"),
    "vertical_rates": ("VERTICAL RATES", "V RATES", "VRATES"),
    "arbs": ("ORIGIN ARBS", "O.ARBS", "ARBS"),
    "cmdt_notes": ("CMDT NOTE", "SRCHG", "SUR", "SURCHARGE"),
    "special_notes": ("SPECIAL NOTE", "SPCL NOTE"),
    "route_notes": ("ROUTE NOTE", "RN", "RNT", "ROUTE"),
    "freetime": ("FREETIME", "DEMDET", "FREE TIME"),
}
READERS = {
    "rates": read_rates_sheet,
    "rates_port_port": read_rates_sheet,
    "vertical_rates": read_vertical_rates_sheet,
    "arbs": read_arbs_sheet,
    "cmdt_notes": read_cmdt_note_sheet,
    "special_notes": read_special_note_sheet,
    "route_notes": read_route_note_sheet,
    "freetime": read_freetime_sheet,
}
LABELS = {
    "rates": "RATES",
    "rates_port_port": "RATES PORT-PORT",
    "vertical_rates": "VERTICAL RATES",
    "arbs": "ORIGIN ARBS",
    "cmdt_notes": "CMDT NOTE",
    "special_notes": "SPECIAL NOTE",
    "route_notes": "ROUTE NOTE",
    "freetime": "FREETIME",
}


@dataclass
class LoadedFiling:
    """What one OPUS workbook holds, by sheet type rather than by name."""

    sheets: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    # {sheet type: the name it was actually found under} - worth showing,
    # since "read your SRCHG sheet as CMDT NOTE" is a claim the user
    # should be able to check.
    found_as: dict[str, str] = field(default_factory=dict)
    unread: list[str] = field(default_factory=list)

    def rows(self, kind: str) -> list[dict[str, Any]]:
        return self.sheets.get(kind, [])

    @property
    def is_empty(self) -> bool:
        return not any(self.sheets.values())


def load_filing(wb: Workbook) -> LoadedFiling:
    """Read every sheet of a workbook this tool understands.

    A sheet matching no alias is listed in `unread` rather than dropped
    silently - a filing carrying something we don't handle should say so.
    """
    out = LoadedFiling()
    claimed: set[str] = set()
    for kind, aliases in SHEET_ALIASES.items():
        for alias in aliases:
            try:
                name = find_sheet(wb, alias)
            except KeyError:
                continue
            if name in claimed:
                continue
            try:
                rows = READERS[kind](wb, name)
            except Exception:
                continue
            claimed.add(name)
            out.sheets[kind] = rows
            out.found_as[kind] = name
            break
    out.unread = [n for n in wb.sheetnames if n not in claimed]
    return out
