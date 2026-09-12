"""What changed between two filings of the same lane.

Compare answers "does this draft match the filing?". This answers a
different question the team asks every week: "what moved since last
time?" - both sides are finished filings, and the interesting output is
the rates that changed, not the rows that disagree.

Routes are matched on the audit's own identity (every location code, the
terms, the transmodes, the vias) with the rates deliberately left OUT of
the key, so a route whose rate moved reports as a CHANGE rather than
disappearing from one side and appearing as new in the other.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Any

from mrg2opus.audit.compare import audit_row_key

RATE_SLOTS = (("rate_20", "20'"), ("rate_40", "40'"), ("rate_40hc", "40'HC"), ("rate_45", "45'"))


@dataclass(frozen=True)
class RateChange:
    route: str
    commodity: str
    size: str
    was: Decimal | None
    now: Decimal | None

    @property
    def difference(self) -> Decimal | None:
        if self.was is None or self.now is None:
            return None
        return self.now - self.was

    @property
    def percent(self) -> float | None:
        d = self.difference
        if d is None or not self.was:
            return None
        return float(d) / float(self.was) * 100.0


@dataclass
class Delta:
    changes: list[RateChange] = field(default_factory=list)
    added: list[dict[str, Any]] = field(default_factory=list)
    removed: list[dict[str, Any]] = field(default_factory=list)
    unchanged: int = 0

    @property
    def total_routes(self) -> int:
        return self.unchanged + len({c.route for c in self.changes}) + len(self.added) + len(self.removed)


def _decimal(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def _route_label(row: dict[str, Any]) -> str:
    return f"{row.get('origin_code') or '?'} → {row.get('destination_code') or '?'}"


def _commodity(row: dict[str, Any]) -> str:
    return str(row.get("commodity_group_description") or row.get("commodity_group_code") or "-")


def compare_filings(before: list[dict[str, Any]], after: list[dict[str, Any]]) -> Delta:
    """`before` and `after` are two filings' RATES rows, oldest first."""
    old = {audit_row_key(r): r for r in before}
    new = {audit_row_key(r): r for r in after}
    out = Delta()

    for key in new.keys() - old.keys():
        out.added.append(new[key])
    for key in old.keys() - new.keys():
        out.removed.append(old[key])

    for key in old.keys() & new.keys():
        was_row, now_row = old[key], new[key]
        moved = False
        for field_name, size in RATE_SLOTS:
            was, now = _decimal(was_row.get(field_name)), _decimal(now_row.get(field_name))
            if was == now:
                continue
            moved = True
            out.changes.append(RateChange(_route_label(now_row), _commodity(now_row), size, was, now))
        if not moved:
            out.unchanged += 1

    out.changes.sort(key=lambda c: (abs(c.difference) if c.difference is not None else 0), reverse=True)
    out.added.sort(key=_route_label)
    out.removed.sort(key=_route_label)
    return out
