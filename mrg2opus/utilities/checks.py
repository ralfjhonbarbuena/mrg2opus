"""What to look at before submitting a filing.

Each check answers something that costs a resubmission when it is wrong,
and each reports its all-clear as loudly as its findings - "no duplicate
filings" is a thing the auditor confirms, not merely the absence of a
warning.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from mrg2opus.audit.compare import find_duplicate_filings
from mrg2opus.location_bank.store import LocationBankStore
from mrg2opus.schema.charge_codes import is_known_charge_code
from mrg2opus.utilities.workbook import LoadedFiling

# Rows OPUS will not take without these. Sequence numbers are NOT here:
# OPUS assigns them, and a filing that leaves them blank is normal.
_REQUIRED = ("origin_code", "destination_code", "prefix", "cgo_type")
_RATE_FIELDS = ("rate_20", "rate_40", "rate_40hc", "rate_45")


@dataclass(frozen=True)
class Finding:
    check: str
    severity: str  # "error" - OPUS rejects it; "warning" - worth a look
    summary: str
    detail: str = ""


@dataclass
class CheckResult:
    name: str
    passed: bool
    headline: str
    findings: list[Finding]


def _rate_rows(filing: LoadedFiling) -> list[dict[str, Any]]:
    return filing.rows("rates") + filing.rows("rates_port_port")


def check_duplicate_filings(filing: LoadedFiling) -> CheckResult:
    """Two rows identical across the audit's own concat. OPUS rejects the
    upload outright, so this is the one to run first."""
    findings = []
    total = 0
    for kind in ("rates", "rates_port_port"):
        rows = filing.rows(kind)
        if not rows:
            continue
        for concat, count in find_duplicate_filings(rows):
            total += 1
            findings.append(Finding(
                "Duplicate filings", "error",
                f"{filing.found_as.get(kind, kind)}: filed {count} times",
                concat,
            ))
    return CheckResult(
        "Duplicate filings", not findings,
        "No route is filed twice" if not findings else f"{total} route(s) filed more than once",
        findings,
    )


def check_location_codes(filing: LoadedFiling, store: LocationBankStore | None = None) -> CheckResult:
    """Every origin and destination against the Location Bank. An
    unrecognized code is not automatically wrong - the bank holds 331
    ports and the world holds more - which is why these are warnings.
    A typo and a genuinely new port look identical from here."""
    store = store or LocationBankStore()
    seen: dict[str, int] = {}
    for row in _rate_rows(filing) + filing.rows("vertical_rates"):
        for field in ("origin_code", "destination_code"):
            raw = row.get(field)
            if not raw:
                continue
            for code in str(raw).split(";"):
                code = code.strip()
                if code:
                    seen[code] = seen.get(code, 0) + 1

    findings = []
    for code, count in sorted(seen.items()):
        if store.get_by_code(code) is None:
            findings.append(Finding(
                "Location codes", "warning",
                f"{code} is not in the Location Bank",
                f"used on {count} row(s)",
            ))
    return CheckResult(
        "Location codes", not findings,
        f"All {len(seen)} codes are known"
        if not findings else f"{len(findings)} of {len(seen)} codes are not in the Location Bank",
        findings,
    )


def check_charge_codes(filing: LoadedFiling) -> CheckResult:
    """The surcharge codes on the note sheets. A code with no name is one
    OPUS has no charge for.

    ROUTE NOTE is deliberately not read here: those rows are header-only
    by definition - charge_seq and code are always 1 and APP, with no
    per-route surcharge breakdown to check. Reading them anyway also went
    wrong in practice, because some lanes' RN sheets are shifted a column
    and the reader picked up the sequence number as a charge code.
    """
    seen: dict[str, int] = {}
    for kind in ("cmdt_notes", "special_notes"):
        for row in filing.rows(kind):
            code = str(row.get("code") or "").strip().upper()
            # APP is the note's own header row, not a charge; anything
            # without a letter in it is not a code at all.
            if not code or code == "APP" or not code.isalpha():
                continue
            seen[code] = seen.get(code, 0) + 1

    findings = [
        Finding("Charge codes", "warning", f"{code} is not a known charge code", f"used {count} time(s)")
        for code, count in sorted(seen.items()) if not is_known_charge_code(code)
    ]
    return CheckResult(
        "Charge codes", not findings,
        f"All {len(seen)} codes are known"
        if not findings else f"{len(findings)} of {len(seen)} codes are unrecognized",
        findings,
    )


def check_required_fields(filing: LoadedFiling) -> CheckResult:
    """Rows missing something a rate row cannot be filed without, and rows
    carrying no rate in any of the four slots."""
    findings = []
    for kind in ("rates", "rates_port_port"):
        rows = filing.rows(kind)
        sheet = filing.found_as.get(kind, kind)
        for i, row in enumerate(rows, start=3):  # sheet row, past the 2-row header
            missing = [f for f in _REQUIRED if not row.get(f)]
            if missing:
                findings.append(Finding(
                    "Required fields", "error",
                    f"{sheet} row {i}: no {', '.join(missing)}",
                    f"{row.get('origin_code') or '?'} to {row.get('destination_code') or '?'}",
                ))
            elif all(row.get(f) in (None, "") for f in _RATE_FIELDS):
                findings.append(Finding(
                    "Required fields", "warning",
                    f"{sheet} row {i}: no rate in any container size",
                    f"{row.get('origin_code')} to {row.get('destination_code')}",
                ))
    return CheckResult(
        "Required fields", not findings,
        "Every row has its codes and a rate" if not findings else f"{len(findings)} row(s) to look at",
        findings,
    )


def run_all(filing: LoadedFiling, store: LocationBankStore | None = None) -> list[CheckResult]:
    return [
        check_duplicate_filings(filing),
        check_required_fields(filing),
        check_location_codes(filing, store),
        check_charge_codes(filing),
    ]
