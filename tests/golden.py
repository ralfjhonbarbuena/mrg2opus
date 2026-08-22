"""Golden-file diff helpers: read the ground-truth OPUS sheets already baked
into the paired sample workbooks and compare them against a parser's output.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import openpyxl

from mrg2opus.schema import opus_columns as cols


def _find_sheet(wb, sheet_name: str) -> str:
    """Sheet naming isn't fully standardized across lanes (e.g. SAF uses
    'OPUS RATES PORT - PORT' with spaces, other lanes use 'OPUS RATES
    PORT-PORT') - match loosely on whitespace/hyphen instead of exact name."""
    if sheet_name in wb.sheetnames:
        return sheet_name
    normalized_target = re.sub(r"[\s-]+", "", sheet_name).upper()
    for name in wb.sheetnames:
        if re.sub(r"[\s-]+", "", name).upper() == normalized_target:
            return name
    raise KeyError(f"No sheet matching {sheet_name!r} in {wb.sheetnames}")


def read_rates_sheet(path: Path, sheet_name: str) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[_find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=3):
        values = [c.value for c in row[: len(cols.RATES_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.RATES_ROW_FIELDS, values)))
    wb.close()
    return rows


def read_arbs_sheet(path: Path, sheet_name: str = cols.SHEET_NAME_ARBS) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[_find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.ARBS_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.ARBS_ROW_FIELDS, values)))
    wb.close()
    return rows


def read_cmdt_note_sheet(path: Path, sheet_name: str) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[_find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.CMDT_NOTE_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.CMDT_NOTE_ROW_FIELDS, values)))
    wb.close()
    return rows


def read_special_note_sheet(path: Path, sheet_name: str = cols.SHEET_NAME_SPECIAL_NOTE) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[_find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.SPECIAL_NOTE_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.SPECIAL_NOTE_ROW_FIELDS, values)))
    wb.close()
    return rows


def _normalize(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        return int(value) if value.isdigit() else value
    return value


def _normalize_cmdt_value(value: Any) -> Any:
    if hasattr(value, "date") and callable(getattr(value, "date")):
        return value.date()
    if isinstance(value, str):
        value = value.strip()
        if value.isdigit():
            return int(value)
        return value
    return value


def rates_row_key(row: dict[str, Any]) -> tuple:
    # o_via_code disambiguates e.g. "Ganzhou via Shekou" vs "Ganzhou via
    # Yantian" - both resolve to the same origin_code (CNGAN) with
    # different routing and different rates, so it must be part of the key.
    return (
        row.get("origin_code"), row.get("destination_code"), row.get("cgo_type"), row.get("prefix"),
        row.get("o_via_code"), row.get("d_via_code"),
    )


def diff_rates(generated: list[dict[str, Any]], expected: list[dict[str, Any]], ignore_fields: set[str] = frozenset()):
    """Row-level (keyed) diff. Returns (matched, missing, extra, field_mismatches)."""
    gen_by_key = {rates_row_key(r): r for r in generated}
    exp_by_key = {rates_row_key(r): r for r in expected}

    gen_keys, exp_keys = set(gen_by_key), set(exp_by_key)
    missing = exp_keys - gen_keys
    extra = gen_keys - exp_keys
    matched_keys = gen_keys & exp_keys

    field_mismatches = []
    matched = 0
    for key in matched_keys:
        g, e = gen_by_key[key], exp_by_key[key]
        row_ok = True
        for field_name in cols.RATES_ROW_FIELDS:
            if field_name in ignore_fields:
                continue
            gv, ev = _normalize(g.get(field_name)), _normalize(e.get(field_name))
            if gv != ev:
                field_mismatches.append((key, field_name, gv, ev))
                row_ok = False
        if row_ok:
            matched += 1

    return matched, missing, extra, field_mismatches
