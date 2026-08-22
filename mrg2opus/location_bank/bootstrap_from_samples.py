"""One-time miner: read Origin/Destination Code+Description pairs out of the
'*PORT-PORT*' OPUS sheets in the paired ground-truth sample workbooks, and
load them into the Location Bank as source="sample_mined". This is one of
the two bootstrap sources (the other is bootstrap_unlocode.py); sample_mined
records only win when UN/LOCODE doesn't already define the same code
(see LocationBankStore.upsert_location precedence).

Run as: python -m mrg2opus.location_bank.bootstrap_from_samples
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from mrg2opus.location_bank.known_aliases import KNOWN_ALIASES
from mrg2opus.location_bank.models import LocationRecord
from mrg2opus.location_bank.store import LocationBankStore

SAMPLES_DIR = Path(__file__).resolve().parents[2] / "Sample MRGs with OPUS FORMATS"

# Origin Code/Description and Destination Code/Description columns are fixed
# by the shared OPUS RATES(-*) header layout (see schema/opus_columns.py).
ORIGIN_CODE_COL, ORIGIN_DESC_COL = 8, 9
DEST_CODE_COL, DEST_DESC_COL = 14, 15

PORT_PORT_SHEET_RE = re.compile(r"PORT\s*-\s*PORT", re.IGNORECASE)


def _country_from_code(code: str) -> str | None:
    return code[:2] if len(code) >= 2 and code[:2].isalpha() else None


def _scan_workbook(path: Path) -> tuple[dict[str, str], dict[str, set[str]]]:
    """Returns (clean, grouped):
    - clean: code -> name, mined from rows where the PORT-PORT description
      wasn't a combined multi-port string (i.e. this port wasn't grouped
      with others in its raw source row).
    - grouped: description_text -> set of codes seen sharing that exact
      combined description (i.e. siblings from the same raw multi-port
      group), for the elimination pass in bootstrap().
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    clean: dict[str, str] = {}
    grouped: dict[str, set[str]] = {}
    for sheet_name in wb.sheetnames:
        if not PORT_PORT_SHEET_RE.search(sheet_name):
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=3):
            for code_col, desc_col in ((ORIGIN_CODE_COL, ORIGIN_DESC_COL), (DEST_CODE_COL, DEST_DESC_COL)):
                code = row[code_col - 1].value
                desc = row[desc_col - 1].value
                if not code or not desc:
                    continue
                code = str(code).strip()
                desc = str(desc).strip()
                if not code or ";" in code:
                    # PORT-PORT sheets should already be single-code per row;
                    # skip anything unexpectedly grouped rather than mis-mine it.
                    continue
                if ";" in desc:
                    # Known quirk: when a raw row's origin was a multi-port group,
                    # PORT-PORT repeats the WHOLE combined description string on
                    # every exploded sibling row rather than just this port's own
                    # name - collect the sibling code set for the elimination
                    # pass instead of guessing which segment is "this" code's name.
                    grouped.setdefault(desc, set()).add(code)
                elif code not in clean:
                    clean[code] = desc
    wb.close()
    return clean, grouped


def _infer_by_elimination(known: dict[str, str], grouped: dict[str, set[str]]) -> int:
    """A group's combined description is N ';'-joined name segments for its N
    codes, but neither list is in matching order (both are independently
    sorted - see saf.py comments). If N-1 of the N codes already have a known
    name, the last code and the last unmatched name segment must pair up -
    that's a logical deduction, not a guess, so it's safe to mine. Runs to a
    fixed point since resolving one group can unlock another. Returns the
    number of newly inferred entries.
    """
    total_new = 0
    while True:
        new_this_pass = 0
        for desc, codes in grouped.items():
            parts = desc.split(";")
            if len(parts) != len(codes):
                continue  # can't reliably pair up - skip rather than guess
            known_names = {known[c] for c in codes if c in known}
            remaining_codes = [c for c in codes if c not in known]
            remaining_names = [p for p in parts if p not in known_names]
            if len(remaining_codes) == 1 and len(remaining_names) == 1:
                known[remaining_codes[0]] = remaining_names[0]
                new_this_pass += 1
        total_new += new_this_pass
        if new_this_pass == 0:
            break
    return total_new


def bootstrap(store: LocationBankStore | None = None) -> int:
    store = store or LocationBankStore()

    known: dict[str, str] = {}
    grouped: dict[str, set[str]] = {}
    for path in sorted(SAMPLES_DIR.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        clean, file_grouped = _scan_workbook(path)
        for code, name in clean.items():
            known.setdefault(code, name)
        for desc, codes in file_grouped.items():
            grouped.setdefault(desc, set()).update(codes)

    inferred = _infer_by_elimination(known, grouped)
    if inferred:
        print(f"Inferred {inferred} additional location(s) by elimination from grouped sightings")

    for code, name in known.items():
        store.upsert_location(
            LocationRecord(code=code, primary_name=name, country=_country_from_code(code), source="sample_mined")
        )

    for alias, code in KNOWN_ALIASES.items():
        if code in known:
            store.add_alias(alias, code, source="manual_override")

    return len(known)


if __name__ == "__main__":
    n = bootstrap()
    print(f"Mined and upserted {n} location records from {SAMPLES_DIR}")
