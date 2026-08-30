"""Merge multiple dated raw-MRG snapshots of the same TAD scope into one
continuous filing history.

Confirmed against reference/1_MRGs/23_TAD FILING AEW AMW's own ground
truth: the folder ships TWO raw "AEW"/"AMW" files for the same nominal
period ("Sep MRG dated 20th Aug", validity 2026-09-01 to 2026-09-15, and
"Sep MRG dated 27th Aug", validity 2026-09-07 to 2026-09-15 at revised
rates) - a rate correction issued mid-period, not a "pick one" choice.
Ground truth files BOTH as separate CMDT NOTE blocks: the earlier file's
validity is truncated to end the day before the later file's own validity
starts (2026-09-01 to 2026-09-06), and the later file's own window is used
as-is. Every row within one raw file shares one uniform validity window
(confirmed both scopes, both files) - so this operates on whole snapshot
occurrences, not per-row.

The user asked for this to be a general TAD mechanism, not AEW/AMW-specific
- OEW/OMW and WMW/WEW route through it too (see their own parse_raw()),
though with a single occurrence today it's a no-op.
"""
from __future__ import annotations

import re
from datetime import date, timedelta
from typing import Protocol, TypeVar

from openpyxl.workbook import Workbook


class _HasValidity(Protocol):
    validity_start: date | None
    validity_end: date | None


RawRowT = TypeVar("RawRowT", bound=_HasValidity)


def merge_dated_snapshots(occurrences: list[list[RawRowT]]) -> list[RawRowT]:
    """occurrences: one row list per raw sheet occurrence (e.g. one per
    uploaded file's own "AEW" sheet - see excel_io/merge.py's TAD sheet-
    collision handling for how multiple same-named sheets survive into one
    workbook). Returns every occurrence's rows concatenated, with all but
    the chronologically-last occurrence's validity_end truncated to the
    day before the next occurrence's validity_start. Rows are mutated
    in place (dataclasses, cheap) - safe since callers don't reuse the
    original per-occurrence lists afterward."""
    non_empty = [occ for occ in occurrences if occ]
    if len(non_empty) <= 1:
        return non_empty[0] if non_empty else []

    non_empty.sort(key=lambda occ: occ[0].validity_start or date.min)
    merged: list[RawRowT] = []
    for i, occ in enumerate(non_empty):
        if i + 1 < len(non_empty):
            next_start = non_empty[i + 1][0].validity_start
            cutoff = (next_start - timedelta(days=1)) if next_start else None
            for row in occ:
                if cutoff is not None and (row.validity_end is None or row.validity_end > cutoff):
                    row.validity_end = cutoff
                merged.append(row)
        else:
            merged.extend(occ)
    return merged


def find_snapshot_sheets(wb: Workbook, base_name: str) -> list[str]:
    """A second (or third...) dated snapshot of the same base sheet name
    survives the upload merge (excel_io/merge.py) renamed "{base} (2)",
    "{base} (3)", etc. rather than raising DuplicateSheetError - find all
    of them, in that numeric order (the base name itself first)."""
    pattern = re.compile(rf"^{re.escape(base_name)}(?: \((\d+)\))?$")
    matches = [(m.group(1), name) for name in wb.sheetnames if (m := pattern.match(name))]
    matches.sort(key=lambda pair: int(pair[0]) if pair[0] else 1)
    return [name for _, name in matches]
