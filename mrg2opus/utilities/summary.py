"""What is in this filing - the screen you want before opening the file.

Answers the questions asked of an unfamiliar workbook in that order: what
sheets, how many rows, which commodity groups, which cargo types, how
many ports, and what dates it covers.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

from mrg2opus.utilities.workbook import LABELS, LoadedFiling


@dataclass
class FilingSummary:
    sheets: list[dict[str, Any]] = field(default_factory=list)
    groups: list[dict[str, Any]] = field(default_factory=list)
    cargo_types: list[dict[str, Any]] = field(default_factory=list)
    origins: int = 0
    destinations: int = 0
    rate_rows: int = 0
    validity: tuple[date | None, date | None] = (None, None)
    unread_sheets: list[str] = field(default_factory=list)


def _codes(rows: list[dict[str, Any]], field_name: str) -> set[str]:
    """Distinct ports, counting a ";"-joined group as the ports in it."""
    out: set[str] = set()
    for row in rows:
        raw = row.get(field_name)
        if not raw:
            continue
        out.update(c.strip() for c in str(raw).split(";") if c.strip())
    return out


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    return value if isinstance(value, date) else None


def _validity(filing: LoadedFiling) -> tuple[date | None, date | None]:
    """The window the filing covers, off the note sheets' own application
    dates - the only place in an OPUS workbook that states it."""
    starts, ends = [], []
    for kind in ("cmdt_notes", "special_notes", "route_notes"):
        for row in filing.rows(kind):
            if (d := _as_date(row.get("application_effective"))):
                starts.append(d)
            if (d := _as_date(row.get("application_expires"))):
                ends.append(d)
    return (min(starts) if starts else None, max(ends) if ends else None)


def summarize(filing: LoadedFiling) -> FilingSummary:
    rates = filing.rows("rates") or filing.rows("rates_port_port")
    out = FilingSummary(
        rate_rows=len(rates),
        origins=len(_codes(rates, "origin_code")),
        destinations=len(_codes(rates, "destination_code")),
        validity=_validity(filing),
        unread_sheets=list(filing.unread),
    )
    out.sheets = [
        {"Sheet type": LABELS[kind], "Named": filing.found_as[kind], "Rows": len(filing.rows(kind))}
        for kind in LABELS if kind in filing.found_as
    ]

    by_group = Counter(
        (str(r.get("commodity_group_code") or "-"), str(r.get("commodity_group_description") or "-"))
        for r in rates
    )
    out.groups = [
        {"Code": code, "Description": desc, "Rows": n}
        for (code, desc), n in sorted(by_group.items(), key=lambda kv: -kv[1])
    ]

    by_type = Counter((str(r.get("prefix") or "-"), str(r.get("cgo_type") or "-")) for r in rates)
    out.cargo_types = [
        {"Prefix": prefix, "CGO Type": cgo, "Rows": n}
        for (prefix, cgo), n in sorted(by_type.items(), key=lambda kv: -kv[1])
    ]
    return out
