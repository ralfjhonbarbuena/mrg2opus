# MRG-vs-OPUS Comparison Procedure Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a user upload raw MRG file(s) plus one reference OPUS-format Excel file, in a new standalone "Compare" mode (separate from the existing 5-file-then-4-step wizard), and see a row-level and field-level diff between what mrg2opus would generate and what the reference file actually contains, across all 5 OPUS sheet types.

**Architecture:** A new production module `mrg2opus/audit/compare.py` holds sheet readers and two diff strategies — a generalized keyed diff (`diff_by_key`, reused for RATES/RATES PORT-PORT/ARBS) and a block-reconstruction diff (`diff_cmdt_blocks`, for CMDT NOTE/SPECIAL NOTE, which have no reliable per-row key). `tests/golden.py`'s existing reader/diff functions become thin wrappers delegating to this module instead of duplicating its logic. A new `mrg2opus/ui/mrg_upload.py` factors the upload-fingerprint-merge-classify logic shared by the wizard's Step 1 and the new Compare page. A new `mrg2opus/ui/compare_page.py` renders the Compare flow, reached via a mode toggle added to `app.py`.

**Tech Stack:** Python 3.14, pydantic v2, openpyxl, Streamlit, pytest. No new dependencies.

**Spec:** `docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md`

## Global Constraints

- Every existing test in `tests/` must keep passing unchanged after this plan — the full suite currently passes at 54/54 (`./.venv/Scripts/python.exe -m pytest tests/ -q`). Task 3 in particular touches `tests/golden.py`, used by every lane's golden-diff tests; verify the full suite immediately after that task, not just the new tests.
- No CLI subcommand for Compare (spec's explicit non-goal) — UI only.
- `RatesPortPortRow` is always a deterministic transform of `RatesRow` (`explode_rates_row()`) — the "Generate MRG as: Grouped/Exploded/Both" selector controls which derived form(s) get compared, never a second parse path.
- This plan implements ONLY the Compare spec. The separately-approved Phase 3 Audit Gate spec (`docs/superpowers/specs/2026-08-23-phase3-audit-gate-design.md`, internal self-check rules + wizard Step 4) is out of scope here and gets its own plan later — `app.py`'s existing 4-step `STEP_LABELS`/`STEP_RENDERERS` are not renumbered by this plan.
- Follow the project's existing git identity convention: commits use `git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit ...` (no global git config is set on this machine, and it must not be changed).

---

## Task 1: `audit/compare.py` — sheet readers and keyed diff (RATES / RATES PORT-PORT / ARBS)

**Files:**
- Create: `mrg2opus/audit/compare.py`
- Test: `tests/test_audit_compare.py`

**Interfaces:**
- Consumes: `mrg2opus.schema.opus_columns` (`RATES_ROW_FIELDS`, `ARBS_ROW_FIELDS`, `SHEET_NAME_RATES`, `SHEET_NAME_ARBS`), `openpyxl.workbook.Workbook`.
- Produces (used by later tasks):
  - `find_sheet(wb: Workbook, sheet_name: str) -> str`
  - `read_rates_sheet(wb: Workbook, sheet_name: str) -> list[dict[str, Any]]`
  - `read_arbs_sheet(wb: Workbook, sheet_name: str = cols.SHEET_NAME_ARBS) -> list[dict[str, Any]]`
  - `rates_row_key(row: dict[str, Any]) -> tuple`
  - `arbs_row_key(row: dict[str, Any]) -> tuple`
  - `@dataclass KeyedDiffResult(matched: int, missing: set[tuple], extra: set[tuple], field_mismatches: list[tuple[tuple, str, Any, Any]])`
  - `diff_by_key(generated, expected, key_fn, fields, ignore_fields=frozenset()) -> KeyedDiffResult`
  - `_normalize(value: Any) -> Any` (module-private, also used by Task 2)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_audit_compare.py`:

```python
from __future__ import annotations

import openpyxl

from mrg2opus.audit.compare import (
    arbs_row_key,
    diff_by_key,
    find_sheet,
    rates_row_key,
    read_arbs_sheet,
    read_rates_sheet,
)
from mrg2opus.schema import opus_columns as cols


def _rates_sheet_wb(rows: list[list]) -> "openpyxl.Workbook":
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cols.SHEET_NAME_RATES
    # Rows start at row 3 (rows 1-2 are the 2-row header) - matches every
    # lane's writer output and golden.py's existing read_rates_sheet.
    for offset, row in enumerate(rows):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=3 + offset, column=col_idx, value=value)
    return wb


def test_find_sheet_exact_match():
    wb = openpyxl.Workbook()
    wb.active.title = "OPUS RATES"
    assert find_sheet(wb, "OPUS RATES") == "OPUS RATES"


def test_find_sheet_loose_whitespace_hyphen_match():
    wb = openpyxl.Workbook()
    wb.active.title = "OPUS RATES PORT - PORT"  # SAF's spacing quirk
    assert find_sheet(wb, "OPUS RATES PORT-PORT") == "OPUS RATES PORT - PORT"


def test_find_sheet_raises_when_missing():
    wb = openpyxl.Workbook()
    wb.active.title = "Something Else"
    try:
        find_sheet(wb, "OPUS RATES")
        assert False, "expected KeyError"
    except KeyError:
        pass


def test_read_rates_sheet_skips_blank_rows_and_reads_by_field_order():
    row = [None] * len(cols.RATES_ROW_FIELDS)
    row[cols.RATES_ROW_FIELDS.index("origin_code")] = "CNSHA"
    row[cols.RATES_ROW_FIELDS.index("destination_code")] = "USLAX"
    wb = _rates_sheet_wb([row, [None] * len(cols.RATES_ROW_FIELDS)])
    rows = read_rates_sheet(wb, cols.SHEET_NAME_RATES)
    assert len(rows) == 1
    assert rows[0]["origin_code"] == "CNSHA"
    assert rows[0]["destination_code"] == "USLAX"


def test_arbs_row_key():
    assert arbs_row_key({"point": "CNSHA", "over": "20MT", "per": "MT", "seq": 1}) == ("CNSHA", "20MT", "MT")


def test_diff_by_key_matched_missing_extra_and_field_mismatch():
    generated = [
        {"k": "A", "v": 1},
        {"k": "B", "v": 2},
        {"k": "C", "v": 99},
    ]
    expected = [
        {"k": "A", "v": 1},
        {"k": "C", "v": 3},
        {"k": "D", "v": 4},
    ]
    result = diff_by_key(generated, expected, key_fn=lambda r: (r["k"],), fields=["v"])
    assert result.matched == 1
    assert result.missing == {("D",)}
    assert result.extra == {("B",)}
    assert result.field_mismatches == [(("C",), "v", 99, 3)]


def test_diff_by_key_respects_ignore_fields():
    generated = [{"k": "A", "v": 1, "note": "x"}]
    expected = [{"k": "A", "v": 1, "note": "y"}]
    result = diff_by_key(
        generated, expected, key_fn=lambda r: (r["k"],), fields=["v", "note"], ignore_fields={"note"}
    )
    assert result.matched == 1
    assert result.field_mismatches == []
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_audit_compare.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'mrg2opus.audit.compare'` (the module doesn't exist yet).

- [ ] **Step 3: Write the implementation**

Create `mrg2opus/audit/compare.py`:

```python
"""Compare a parsed MRG's OpusRowSet against a reference OPUS-format Excel
file, sheet by sheet. This is the production home of the sheet-reading
and diff logic tests/golden.py has always used for golden-file testing
against the 5 bundled samples - both the tests and the UI's Compare mode
call this module, so there is one implementation instead of two that can
drift apart.

See docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Callable

from openpyxl.workbook import Workbook

from mrg2opus.schema import opus_columns as cols


def find_sheet(wb: Workbook, sheet_name: str) -> str:
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


def read_rates_sheet(wb: Workbook, sheet_name: str) -> list[dict[str, Any]]:
    ws = wb[find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=3):
        values = [c.value for c in row[: len(cols.RATES_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.RATES_ROW_FIELDS, values)))
    return rows


def read_arbs_sheet(wb: Workbook, sheet_name: str = cols.SHEET_NAME_ARBS) -> list[dict[str, Any]]:
    ws = wb[find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.ARBS_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.ARBS_ROW_FIELDS, values)))
    return rows


def read_cmdt_note_sheet(wb: Workbook, sheet_name: str) -> list[dict[str, Any]]:
    ws = wb[find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.CMDT_NOTE_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.CMDT_NOTE_ROW_FIELDS, values)))
    return rows


def read_special_note_sheet(wb: Workbook, sheet_name: str = cols.SHEET_NAME_SPECIAL_NOTE) -> list[dict[str, Any]]:
    ws = wb[find_sheet(wb, sheet_name)]
    rows = []
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row[: len(cols.SPECIAL_NOTE_ROW_FIELDS)]]
        if all(v is None for v in values):
            continue
        rows.append(dict(zip(cols.SPECIAL_NOTE_ROW_FIELDS, values)))
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


def rates_row_key(row: dict[str, Any]) -> tuple:
    # o_via_code disambiguates e.g. "Ganzhou via Shekou" vs "Ganzhou via
    # Yantian" - both resolve to the same origin_code (CNGAN) with
    # different routing and different rates, so it must be part of the key.
    return (
        row.get("origin_code"), row.get("destination_code"), row.get("cgo_type"), row.get("prefix"),
        row.get("o_via_code"), row.get("d_via_code"),
    )


def arbs_row_key(row: dict[str, Any]) -> tuple:
    return (row.get("point"), row.get("over"), row.get("per"))


@dataclass
class KeyedDiffResult:
    matched: int
    missing: set[tuple]
    extra: set[tuple]
    field_mismatches: list[tuple[tuple, str, Any, Any]] = field(default_factory=list)


def diff_by_key(
    generated: list[dict[str, Any]],
    expected: list[dict[str, Any]],
    key_fn: Callable[[dict[str, Any]], tuple],
    fields: list[str],
    ignore_fields: set[str] = frozenset(),
) -> KeyedDiffResult:
    """Row-level (keyed) diff, generalized from the matching logic every
    lane's golden test already used via rates_row_key specifically."""
    gen_by_key = {key_fn(r): r for r in generated}
    exp_by_key = {key_fn(r): r for r in expected}

    gen_keys, exp_keys = set(gen_by_key), set(exp_by_key)
    missing = exp_keys - gen_keys
    extra = gen_keys - exp_keys
    matched_keys = gen_keys & exp_keys

    field_mismatches: list[tuple[tuple, str, Any, Any]] = []
    matched = 0
    for key in matched_keys:
        g, e = gen_by_key[key], exp_by_key[key]
        row_ok = True
        for field_name in fields:
            if field_name in ignore_fields:
                continue
            gv, ev = _normalize(g.get(field_name)), _normalize(e.get(field_name))
            if gv != ev:
                field_mismatches.append((key, field_name, gv, ev))
                row_ok = False
        if row_ok:
            matched += 1

    return KeyedDiffResult(matched=matched, missing=missing, extra=extra, field_mismatches=field_mismatches)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_audit_compare.py -v`
Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add mrg2opus/audit/compare.py tests/test_audit_compare.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Add audit/compare.py: sheet readers and keyed diff for RATES/ARBS"
```

---

## Task 2: `audit/compare.py` — block-based diff for CMDT NOTE / SPECIAL NOTE

**Files:**
- Modify: `mrg2opus/audit/compare.py` (append)
- Test: `tests/test_audit_compare.py` (append)

**Interfaces:**
- Consumes: `_normalize` from Task 1 (same module).
- Produces (used by Task 4 and Task 6):
  - `@dataclass CmdtBlock(key: str, parent: dict, children: list[dict])`
  - `reconstruct_blocks(rows: list[dict[str, Any]]) -> list[CmdtBlock]`
  - `@dataclass BlockDiffResult(missing_blocks: list[str], extra_blocks: list[str], field_mismatches: list[tuple[str, int, str, Any, Any]])`
  - `diff_cmdt_blocks(generated: list[dict], expected: list[dict], fields: list[str]) -> BlockDiffResult`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_audit_compare.py`:

```python
from mrg2opus.audit.compare import diff_cmdt_blocks, reconstruct_blocks


def _cmdt_row(header_seq=None, note_seq=None, contents=None, charge_seq=None, code=None, amount=None) -> dict:
    row = dict.fromkeys(cols.CMDT_NOTE_ROW_FIELDS)
    row.update(header_seq=header_seq, note_seq=note_seq, contents=contents, charge_seq=charge_seq, code=code, amount=amount)
    return row


def test_reconstruct_blocks_groups_children_under_parent():
    rows = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
        _cmdt_row(header_seq=2, note_seq=2, contents="Block B", charge_seq=1, code="EFS"),
    ]
    blocks = reconstruct_blocks(rows)
    assert [b.key for b in blocks] == ["Block A", "Block B"]
    assert len(blocks[0].children) == 1
    assert blocks[0].children[0]["code"] == "OBS"
    assert len(blocks[1].children) == 0


def test_diff_cmdt_blocks_matched_identical():
    generated = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
    ]
    expected = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
    ]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == []
    assert result.extra_blocks == []
    assert result.field_mismatches == []


def test_diff_cmdt_blocks_missing_and_extra():
    generated = [_cmdt_row(header_seq=1, note_seq=1, contents="Only Generated", charge_seq=1, code="PSS")]
    expected = [_cmdt_row(header_seq=1, note_seq=1, contents="Only Reference", charge_seq=1, code="PSS")]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == ["Only Reference"]
    assert result.extra_blocks == ["Only Generated"]


def test_diff_cmdt_blocks_field_mismatch_within_matched_block():
    generated = [_cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS", amount=100)]
    expected = [_cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS", amount=200)]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == []
    assert result.extra_blocks == []
    assert len(result.field_mismatches) == 1
    key, idx, field_name, gv, ev = result.field_mismatches[0]
    assert (key, idx, field_name, gv, ev) == ("Block A", 0, "amount", 100, 200)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_audit_compare.py -v`
Expected: FAIL with `ImportError: cannot import name 'diff_cmdt_blocks'` on the 4 new tests; the 7 Task 1 tests still pass.

- [ ] **Step 3: Write the implementation**

Append to `mrg2opus/audit/compare.py`:

```python
@dataclass
class CmdtBlock:
    key: str
    parent: dict[str, Any]
    children: list[dict[str, Any]] = field(default_factory=list)


def reconstruct_blocks(rows: list[dict[str, Any]]) -> list[CmdtBlock]:
    """A row with non-None header_seq OR note_seq starts a new block
    (mirrors the writer's own fill-down convention, see
    excel_io/writer.py); every following row until the next such marker
    belongs to it. Blocks are keyed by the parent's `contents` text - the
    block's human-readable identity, stable regardless of row order."""
    blocks: list[CmdtBlock] = []
    current: CmdtBlock | None = None
    for row in rows:
        is_parent = row.get("header_seq") is not None or row.get("note_seq") is not None
        if is_parent:
            key = str(row.get("contents") or "").strip()
            current = CmdtBlock(key=key, parent=row, children=[])
            blocks.append(current)
        elif current is not None:
            current.children.append(row)
    return blocks


@dataclass
class BlockDiffResult:
    missing_blocks: list[str] = field(default_factory=list)
    extra_blocks: list[str] = field(default_factory=list)
    field_mismatches: list[tuple[str, int, str, Any, Any]] = field(default_factory=list)


def diff_cmdt_blocks(generated: list[dict[str, Any]], expected: list[dict[str, Any]], fields: list[str]) -> BlockDiffResult:
    """CMDT NOTE / SPECIAL NOTE have no reliable per-row key (children
    share their parent's blank header_seq/note_seq) and a reference
    file's row order isn't guaranteed to match the generator's - unlike
    golden tests, which only ever compare against one exact known sample
    and can safely assume matching order. Blocks are matched by
    contents-text key; for matched blocks, child rows are compared
    positionally - the fragile ordering assumption is scoped to one
    block's internal order, not the whole sheet."""
    gen_blocks = {b.key: b for b in reconstruct_blocks(generated)}
    exp_blocks = {b.key: b for b in reconstruct_blocks(expected)}

    missing_blocks = sorted(set(exp_blocks) - set(gen_blocks))
    extra_blocks = sorted(set(gen_blocks) - set(exp_blocks))

    field_mismatches: list[tuple[str, int, str, Any, Any]] = []
    for key in set(gen_blocks) & set(exp_blocks):
        g_block, e_block = gen_blocks[key], exp_blocks[key]
        g_rows = [g_block.parent, *g_block.children]
        e_rows = [e_block.parent, *e_block.children]
        for idx, (g_row, e_row) in enumerate(zip(g_rows, e_rows)):
            for field_name in fields:
                gv, ev = _normalize(g_row.get(field_name)), _normalize(e_row.get(field_name))
                if gv != ev:
                    field_mismatches.append((key, idx, field_name, gv, ev))
        if len(g_rows) != len(e_rows):
            field_mismatches.append((key, -1, "_row_count", len(g_rows), len(e_rows)))

    return BlockDiffResult(missing_blocks=missing_blocks, extra_blocks=extra_blocks, field_mismatches=field_mismatches)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_audit_compare.py -v`
Expected: PASS, 11 tests.

- [ ] **Step 5: Commit**

```bash
git add mrg2opus/audit/compare.py tests/test_audit_compare.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Add block-based diff for CMDT NOTE / SPECIAL NOTE"
```

---

## Task 3: Refactor `tests/golden.py` to delegate to `audit/compare.py`

**Files:**
- Modify: `tests/golden.py` (full rewrite of the parts described below)

**Interfaces:**
- Consumes: `find_sheet`, `read_rates_sheet`, `read_arbs_sheet`, `read_cmdt_note_sheet`, `read_special_note_sheet`, `rates_row_key`, `diff_by_key` from `mrg2opus.audit.compare` (Task 1).
- Produces: **unchanged public API** — every existing test file imports `_normalize`, `_normalize_cmdt_value`, `diff_rates`, `read_rates_sheet`, `read_arbs_sheet`, `read_cmdt_note_sheet`, `read_special_note_sheet` from `tests.golden` with their current path-based signatures and behavior. No test file outside `tests/golden.py` itself changes in this task.

This task has no new tests of its own — its correctness is proven by every existing golden test in `tests/test_parsers_*.py` continuing to pass unchanged. Confirmed via grep before writing this plan: `_normalize` is imported directly by `test_parsers_cse.py`, `test_parsers_laec.py`, `test_parsers_lawc.py`; `_normalize_cmdt_value` is imported directly by `test_parsers_eaf.py` and `test_parsers_saf.py` — both stay, unchanged, verbatim, in `golden.py`. Every `read_*_sheet(path, sheet_name)` and `diff_rates(generated, expected, ignore_fields=...)` call site across all 5 lane test files passes a file **path** as the first argument and unpacks `diff_rates`'s result as a 4-tuple — both of those exact calling conventions are preserved.

- [ ] **Step 1: Replace `tests/golden.py`'s content**

Replace the entire file with:

```python
"""Golden-file diff helpers: read the ground-truth OPUS sheets already baked
into the paired sample workbooks and compare them against a parser's output.

Sheet-reading and keyed-diff logic live in mrg2opus.audit.compare (the
production module the UI's Compare mode also calls) - this module is a
thin path-loading wrapper so every existing test call site here keeps its
current path-based signature and (matched, missing, extra, mismatches)
4-tuple return shape unchanged.
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import openpyxl

from mrg2opus.audit.compare import diff_by_key, rates_row_key
from mrg2opus.audit.compare import read_arbs_sheet as _read_arbs_sheet_from_wb
from mrg2opus.audit.compare import read_cmdt_note_sheet as _read_cmdt_note_sheet_from_wb
from mrg2opus.audit.compare import read_rates_sheet as _read_rates_sheet_from_wb
from mrg2opus.audit.compare import read_special_note_sheet as _read_special_note_sheet_from_wb
from mrg2opus.schema import opus_columns as cols


def read_rates_sheet(path: Path, sheet_name: str) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return _read_rates_sheet_from_wb(wb, sheet_name)
    finally:
        wb.close()


def read_arbs_sheet(path: Path, sheet_name: str = cols.SHEET_NAME_ARBS) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return _read_arbs_sheet_from_wb(wb, sheet_name)
    finally:
        wb.close()


def read_cmdt_note_sheet(path: Path, sheet_name: str) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return _read_cmdt_note_sheet_from_wb(wb, sheet_name)
    finally:
        wb.close()


def read_special_note_sheet(path: Path, sheet_name: str = cols.SHEET_NAME_SPECIAL_NOTE) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return _read_special_note_sheet_from_wb(wb, sheet_name)
    finally:
        wb.close()


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


def diff_rates(generated: list[dict[str, Any]], expected: list[dict[str, Any]], ignore_fields: set[str] = frozenset()):
    """Row-level (keyed) diff. Returns (matched, missing, extra, field_mismatches) -
    unchanged shape from before this refactor, so every existing call
    site's 4-way unpack keeps working."""
    result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS, ignore_fields=ignore_fields)
    return result.matched, result.missing, result.extra, result.field_mismatches
```

- [ ] **Step 2: Run the full test suite to verify nothing broke**

Run: `./.venv/Scripts/python.exe -m pytest tests/ -v`
Expected: PASS, same 54 tests as before this task (plus the 11 from Tasks 1-2 = 65 total). If any `test_parsers_*.py` test fails, stop and diagnose before continuing — this task must be a pure refactor with zero behavior change.

- [ ] **Step 3: Commit**

```bash
git add tests/golden.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Refactor tests/golden.py to delegate to audit/compare.py"
```

---

## Task 4: Regression test — new compare engine against all 5 bundled samples

**Files:**
- Create: `tests/test_compare_engine_regression.py`

**Interfaces:**
- Consumes: `diff_by_key`, `rates_row_key`, `read_rates_sheet` from `mrg2opus.audit.compare` (Task 1); `SAFParser`, `EAFParser`, `CSEParser`, `LAECParser`, `LAWCParser` (`.run_multi(wb, config) -> dict[str, OpusRowSet]`, confirmed in `mrg2opus/parsers/base.py:51` — every parser has this method, default-lane parsers return `{"": OpusRowSet}`).

- [ ] **Step 1: Write the test**

Create `tests/test_compare_engine_regression.py`:

```python
"""Regression check: run the NEW production compare engine (audit/compare.py,
not tests/golden.py's own helpers) against each bundled sample's own
ground-truth RATES sheet. Every row that should exist must exist on both
sides - the same row-level guarantee test_parsers_*.py already asserts
via `assert not missing` / `assert not extra`, re-derived through the new
production code path as a consistency check on top of what those tests
already cover.

Deliberately does NOT assert zero field_mismatches - several lanes have
documented, accepted field-level gaps (see each lane's own test file
comments, e.g. test_parsers_cse.py's PAMIT rate_20 note). This test
checks row existence, not the field-level nuances the existing
test_parsers_*.py files already own and assert on directly.
"""
from __future__ import annotations

import openpyxl
import pytest

from mrg2opus.audit.compare import diff_by_key, rates_row_key, read_rates_sheet
from mrg2opus.parsers.cse import CSEParser
from mrg2opus.parsers.eaf import EAFParser
from mrg2opus.parsers.laec import LAECParser
from mrg2opus.parsers.lawc import LAWCParser
from mrg2opus.parsers.saf import SAFParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols


@pytest.mark.parametrize(
    "path,parser_cls",
    [
        ("Sample MRGs with OPUS FORMATS/SAF.xlsx", SAFParser),
        ("Sample MRGs with OPUS FORMATS/EAF.xlsx", EAFParser),
        ("Sample MRGs with OPUS FORMATS/CSE.xlsx", CSEParser),
        ("Sample MRGs with OPUS FORMATS/LAEC.xlsx", LAECParser),
        ("Sample MRGs with OPUS FORMATS/LAWC.xlsx", LAWCParser),
    ],
)
def test_compare_engine_finds_no_missing_or_extra_rates_rows_against_own_ground_truth(path, parser_cls):
    wb = openpyxl.load_workbook(path, data_only=True)
    parser = parser_cls()
    row_sets = parser.run_multi(wb, MappingProfile())
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        generated = [r.model_dump() for r in row_set.rates]
        expected = read_rates_sheet(wb, f"OPUS RATES{tag}")

        result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS)
        assert not result.missing, f"{path} [{suffix or '(default)'}]: missing {len(result.missing)} rows"
        assert not result.extra, f"{path} [{suffix or '(default)'}]: {len(result.extra)} unexpected rows"
```

- [ ] **Step 2: Run the test to verify it passes**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_compare_engine_regression.py -v`
Expected: PASS, 5 parametrized cases. If any fails, the issue is almost certainly in how the test itself computed the sheet name/suffix tag (compare against `_output_sheet_names()` in `mrg2opus/ui/steps/step3_customize.py` for the exact convention) — the parsers and readers are already proven correct by Tasks 1-3's passing suite, so a failure here means a bug in this test's own sheet-name construction, not the underlying engine.

- [ ] **Step 3: Commit**

```bash
git add tests/test_compare_engine_regression.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Add compare-engine regression test against all 5 bundled samples"
```

---

## Task 5: `ui/mrg_upload.py` — shared upload/classify helper

**Files:**
- Create: `mrg2opus/ui/mrg_upload.py`
- Test: `tests/test_mrg_upload.py`
- Modify: `mrg2opus/ui/steps/step1_upload.py`

**Interfaces:**
- Consumes: `merge_workbooks`, `DuplicateSheetError` from `mrg2opus.excel_io.merge`; `classify_all`, `ClassificationResult` from `mrg2opus.parsers.registry`.
- Produces (used by Task 6):
  - `fingerprint_uploads(names: list[str], payloads: list[bytes]) -> str`
  - `load_and_classify(payloads: list[bytes]) -> tuple[Workbook, list[ClassificationResult]]` (raises `DuplicateSheetError` or the underlying load exception on failure — does not catch)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_mrg_upload.py`:

```python
from __future__ import annotations

from pathlib import Path

import pytest

from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.parsers import cse, eaf, laec, lawc, saf  # noqa: F401 - registers LayoutProfiles
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify


def test_fingerprint_uploads_stable_for_identical_input():
    assert fingerprint_uploads(["a.xlsx"], [b"same bytes"]) == fingerprint_uploads(["a.xlsx"], [b"same bytes"])


def test_fingerprint_uploads_differs_when_same_name_different_bytes():
    assert fingerprint_uploads(["a.xlsx"], [b"version 1"]) != fingerprint_uploads(["a.xlsx"], [b"version 2"])


def test_fingerprint_uploads_differs_on_order():
    assert fingerprint_uploads(["a.xlsx", "b.xlsx"], [b"1", b"2"]) != fingerprint_uploads(["b.xlsx", "a.xlsx"], [b"2", b"1"])


def test_load_and_classify_returns_workbook_and_ranked_results():
    payload = Path("Sample MRGs with OPUS FORMATS/SAF.xlsx").read_bytes()
    wb, results = load_and_classify([payload])
    assert "SAF" in wb.sheetnames
    assert results[0].profile.lane_id == "SAF"


def test_load_and_classify_raises_on_duplicate_sheet_names():
    payload = Path("Sample MRGs with OPUS FORMATS/SAF.xlsx").read_bytes()
    with pytest.raises(DuplicateSheetError):
        load_and_classify([payload, payload])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_mrg_upload.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'mrg2opus.ui.mrg_upload'`.

- [ ] **Step 3: Write the implementation**

Create `mrg2opus/ui/mrg_upload.py`:

```python
"""Shared MRG upload -> classify logic used by both the wizard's Step 1
and Compare mode - factored out so both call one implementation instead
of two copies that can silently drift apart.
"""
from __future__ import annotations

import hashlib
import io

import openpyxl
from openpyxl.workbook import Workbook

from mrg2opus.excel_io.merge import merge_workbooks
from mrg2opus.parsers.registry import ClassificationResult, classify_all


def fingerprint_uploads(names: list[str], payloads: list[bytes]) -> str:
    """sha256 over each file's name AND bytes - re-uploading an edited
    file under the same name must invalidate any cache keyed on this,
    which a names-only comparison would miss (see MIGRATION_NOTES.md's
    note on the wizard's upload cache, section 3.10)."""
    digest = hashlib.sha256()
    for name, payload in zip(names, payloads):
        digest.update(name.encode("utf-8"))
        digest.update(hashlib.sha256(payload).digest())
    return digest.hexdigest()


def load_and_classify(payloads: list[bytes]) -> tuple[Workbook, list[ClassificationResult]]:
    """Load each payload as a workbook, merge them into one (raises
    excel_io.merge.DuplicateSheetError on a repeated sheet name across
    inputs), and classify the result. Raises on failure rather than
    catching - error PRESENTATION (st.error wording/placement) stays a
    caller concern since it differs slightly between the wizard and
    Compare mode."""
    workbooks = [openpyxl.load_workbook(io.BytesIO(payload), data_only=True) for payload in payloads]
    wb = merge_workbooks(workbooks)
    results = classify_all(wb)
    return wb, results
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_mrg_upload.py -v`
Expected: PASS, 5 tests.

- [ ] **Step 5: Refactor `step1_upload.py` to use the shared helper**

Replace the entire content of `mrg2opus/ui/steps/step1_upload.py` with:

```python
from __future__ import annotations

import streamlit as st

from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.parsers.registry import all_profiles
from mrg2opus.presets.models import MappingProfile
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify
from mrg2opus.ui.state import WizardState


def render(state: WizardState) -> None:
    st.subheader("Upload raw MRG Excel file(s)")
    st.caption(
        "Some lanes ship as more than one real-world file - e.g. CSE's main file plus a separate "
        "\"...for VELAG and VEPBL\" file. Upload every file for one filing together; they're merged "
        "into a single workbook before classification."
    )

    uploaded = st.file_uploader("MRG rate sheet(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    if not uploaded:
        st.info("Upload one or more files to classify them and continue.")
        return

    names = [f.name for f in uploaded]
    payloads = [f.getvalue() for f in uploaded]
    upload_key = fingerprint_uploads(names, payloads)

    if upload_key != state.upload_key:
        # New file set (or edited contents) - reset anything downstream so a
        # stale row_sets/profile from a previous upload can't leak into this
        # one. Resetting `profile` here also matters for Step 2's
        # default_commodity_groups snapshot - it must be taken with no
        # overrides carried over from a prior upload, or it wouldn't
        # reflect this file set's true defaults.
        state.upload_names = names
        state.upload_key = upload_key
        state.workbook = None
        state.classification_results = []
        state.selected_lane_id = None
        state.profile = MappingProfile()
        state.row_sets = None
        state.default_commodity_groups = []
        state.output_bytes = None

    if state.workbook is None:
        try:
            state.workbook, state.classification_results = load_and_classify(payloads)
        except DuplicateSheetError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surfaced directly to the user, not swallowed
            st.error(f"Couldn't open one of these as an Excel workbook: {exc}")
            return

    label = state.upload_names[0] if len(state.upload_names) == 1 else f"{len(state.upload_names)} files"
    st.success(f"Loaded **{label}** — sheets: {', '.join(state.workbook.sheetnames)}")

    results = state.classification_results
    best = results[0] if results else None

    st.markdown("#### Classification")
    if best is None:
        st.error("No lane parsers are registered - nothing to classify against.")
        return

    below_threshold = best.confidence < best.profile.min_confidence
    if below_threshold:
        st.warning(
            f"Best match is **{best.profile.lane_id}** at {best.confidence:.0%} confidence, below the "
            f"{best.profile.min_confidence:.0%} threshold. Pick the correct lane manually below."
        )
    else:
        st.metric("Best match", best.profile.lane_id, f"{best.confidence:.0%} confidence")

    with st.expander("Score breakdown (all registered lanes)", expanded=below_threshold):
        st.dataframe(
            [
                {
                    "lane": r.profile.lane_id,
                    "confidence": f"{r.confidence:.0%}",
                    "sheet_name": f"{r.breakdown['sheet_name']:.0%}",
                    "title_keywords": f"{r.breakdown['title_keywords']:.0%}",
                    "header_fingerprint": f"{r.breakdown['header_fingerprint']:.0%}",
                }
                for r in results
            ],
            hide_index=True,
            width="stretch",
        )

    lane_ids = [p.lane_id for p in all_profiles()]
    default_lane = state.selected_lane_id or best.profile.lane_id
    selected = st.selectbox(
        "Lane to use",
        options=lane_ids,
        index=lane_ids.index(default_lane) if default_lane in lane_ids else 0,
        help="Auto-selected from the best classification match; override if it's wrong.",
    )
    state.selected_lane_id = selected

    if st.button("Continue to Preview →", type="primary"):
        state.step = 2
        st.rerun()
```

This is behavior-identical to the version before this task — only the fingerprinting/loading/merging/classifying logic moved into `mrg_upload.py`; every user-facing message, state field, and control flow branch is unchanged.

- [ ] **Step 6: Manually verify the wizard's Step 1 still works**

Start the dev server (`./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py --server.headless true`, or via the project's `.claude/launch.json` "streamlit-ui" config through the Browser preview tool), upload `Sample MRGs with OPUS FORMATS/SAF.xlsx`, and confirm: the success message shows the file name and sheet list, classification shows SAF at high confidence, and "Continue to Preview →" advances to Step 2. This step has no automated equivalent — the project has no existing Streamlit UI test harness, and this task's automated tests already cover the pure logic that changed; this manual check covers the one thing that couldn't be unit-tested, the actual rendered page.

- [ ] **Step 7: Run the full test suite**

Run: `./.venv/Scripts/python.exe -m pytest tests/ -v`
Expected: PASS, all tests from Tasks 1-4 plus this task's 5 new ones, plus the pre-existing suite.

- [ ] **Step 8: Commit**

```bash
git add mrg2opus/ui/mrg_upload.py mrg2opus/ui/steps/step1_upload.py tests/test_mrg_upload.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Factor MRG upload/classify logic into a shared ui/mrg_upload.py helper"
```

---

## Task 6: `ui/compare_page.py` and the mode toggle in `app.py`

**Files:**
- Create: `mrg2opus/ui/compare_page.py`
- Test: `tests/test_compare_page.py`
- Modify: `mrg2opus/ui/app.py`

**Interfaces:**
- Consumes: `arbs_row_key`, `diff_by_key`, `diff_cmdt_blocks`, `rates_row_key`, `read_arbs_sheet`, `read_cmdt_note_sheet`, `read_rates_sheet`, `read_special_note_sheet` from `mrg2opus.audit.compare` (Tasks 1-2); `DuplicateSheetError` from `mrg2opus.excel_io.merge`; `ClassificationResult`, `get_profile` from `mrg2opus.parsers.registry`; `MappingProfile` from `mrg2opus.presets.models`; `fingerprint_uploads`, `load_and_classify` from `mrg2opus.ui.mrg_upload` (Task 5); `run_parser` from `mrg2opus.ui.parsing`.
- Produces: `render() -> None` (called from `app.py`); `_run_comparison(row_sets: dict, ref_wb: Workbook, rates_mode: str) -> list[dict]` (pure, unit-tested directly).

- [ ] **Step 1: Write the failing tests for the pure comparison logic**

Create `tests/test_compare_page.py`:

```python
from __future__ import annotations

import openpyxl

from mrg2opus.schema import opus_columns as cols
from mrg2opus.schema.opus_rows import OpusRowSet, RatesPortPortRow, RatesRow
from mrg2opus.ui.compare_page import _run_comparison


def _rates_row(**overrides) -> RatesRow:
    base = dict(
        commodity_group_code="G0001", commodity_group_description="FAK",
        origin_code="CNSHA", origin_description="Shanghai",
        destination_code="USLAX", destination_description="Los Angeles",
        prefix="D", cgo_type="DR", rate_20=1000,
    )
    base.update(overrides)
    return RatesRow(**base)


def _reference_workbook_with_rates_sheet(rows: list[list]) -> "openpyxl.Workbook":
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cols.SHEET_NAME_RATES
    for offset, row in enumerate(rows):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=3 + offset, column=col_idx, value=value)
    return wb


def test_run_comparison_matches_identical_rates_row():
    generated_row = _rates_row()
    row_sets = {"": OpusRowSet(rates=[generated_row])}

    ref_row = [None] * len(cols.RATES_ROW_FIELDS)
    for field_name, value in generated_row.model_dump().items():
        ref_row[cols.RATES_ROW_FIELDS.index(field_name)] = value
    ref_wb = _reference_workbook_with_rates_sheet([ref_row])

    results = _run_comparison(row_sets, ref_wb, "Both")
    rates_result = next(r for r in results if r["sheet_type"] == "RATES")
    assert rates_result["found_in_reference"] is True
    assert rates_result["matched"] == 1
    assert rates_result["missing"] == []
    assert rates_result["extra"] == []


def test_run_comparison_reports_missing_sheet_as_all_extra():
    generated_row = _rates_row()
    row_sets = {"": OpusRowSet(rates=[generated_row])}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results = _run_comparison(row_sets, ref_wb, "Grouped (RATES)")
    rates_result = next(r for r in results if r["sheet_type"] == "RATES")
    assert rates_result["found_in_reference"] is False
    assert len(rates_result["extra"]) == 1


def test_run_comparison_respects_rates_mode_grouped_only():
    generated_row = _rates_row()
    row_set = OpusRowSet(rates=[generated_row], rates_port_port=[RatesPortPortRow(**generated_row.model_dump())])
    row_sets = {"": row_set}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results = _run_comparison(row_sets, ref_wb, "Grouped (RATES)")
    sheet_types = {r["sheet_type"] for r in results}
    assert sheet_types == {"RATES"}


def test_run_comparison_both_mode_includes_grouped_and_exploded():
    generated_row = _rates_row()
    row_set = OpusRowSet(rates=[generated_row], rates_port_port=[RatesPortPortRow(**generated_row.model_dump())])
    row_sets = {"": row_set}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results = _run_comparison(row_sets, ref_wb, "Both")
    sheet_types = {r["sheet_type"] for r in results}
    assert sheet_types == {"RATES", "RATES PORT-PORT"}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_compare_page.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'mrg2opus.ui.compare_page'`.

- [ ] **Step 3: Write the implementation**

Create `mrg2opus/ui/compare_page.py`:

```python
"""Standalone MRG-vs-reference-OPUS comparison mode - separate from the
5-file-then-4-step wizard. Upload one or more MRG files, upload a
reference OPUS-format Excel file, choose which RATES form(s) to check,
and see where the parser's own output diverges from the reference.

See docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Any

import openpyxl
import streamlit as st
from openpyxl.workbook import Workbook

from mrg2opus.audit.compare import (
    arbs_row_key,
    diff_by_key,
    diff_cmdt_blocks,
    rates_row_key,
    read_arbs_sheet,
    read_cmdt_note_sheet,
    read_rates_sheet,
    read_special_note_sheet,
)
from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.parsers.registry import ClassificationResult, get_profile
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify
from mrg2opus.ui.parsing import run_parser

RATES_MODE_OPTIONS = ["Grouped (RATES)", "Exploded (RATES PORT-PORT)", "Both"]


@dataclass
class CompareState:
    upload_key: str | None = None
    workbook: Workbook | None = None
    classification_results: list[ClassificationResult] = field(default_factory=list)
    selected_lane_id: str | None = None
    reference_workbook: Workbook | None = None
    rates_mode: str = "Both"
    row_sets: dict[str, Any] | None = None
    compare_results: list[dict[str, Any]] | None = None


def _get_state() -> CompareState:
    if "compare" not in st.session_state:
        st.session_state.compare = CompareState()
    return st.session_state.compare


def _compare_keyed_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, key_fn, fields, reader) -> dict:
    sub_lane = suffix or "(default)"
    try:
        expected = reader(ref_wb, sheet_name)
    except KeyError:
        return {
            "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
            "found_in_reference": False, "matched": 0,
            "missing": [], "extra": [{"key": key_fn(r)} for r in generated],
            "field_mismatches": [],
        }
    result = diff_by_key(generated, expected, key_fn=key_fn, fields=fields)
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": result.matched,
        "missing": [{"key": k} for k in result.missing],
        "extra": [{"key": k} for k in result.extra],
        "field_mismatches": [
            {"key": m[0], "field": m[1], "generated": m[2], "reference": m[3]} for m in result.field_mismatches
        ],
    }


def _compare_block_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, fields, reader) -> dict:
    sub_lane = suffix or "(default)"
    try:
        expected = reader(ref_wb, sheet_name)
    except KeyError:
        extra_keys = [
            str(r.get("contents") or "").strip()
            for r in generated
            if r.get("header_seq") is not None or r.get("note_seq") is not None
        ]
        return {
            "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
            "found_in_reference": False, "matched": None,
            "missing": [], "extra": [{"contents": k} for k in extra_keys],
            "field_mismatches": [],
        }
    result = diff_cmdt_blocks(generated, expected, fields)
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": None,
        "missing": [{"contents": k} for k in result.missing_blocks],
        "extra": [{"contents": k} for k in result.extra_blocks],
        "field_mismatches": [
            {"key": m[0], "child_index": m[1], "field": m[2], "generated": m[3], "reference": m[4]}
            for m in result.field_mismatches
        ],
    }


def _run_comparison(row_sets: dict, ref_wb: Workbook, rates_mode: str) -> list[dict]:
    want_grouped = rates_mode in ("Grouped (RATES)", "Both")
    want_exploded = rates_mode in ("Exploded (RATES PORT-PORT)", "Both")

    results: list[dict] = []
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        if want_grouped and row_set.rates:
            results.append(_compare_keyed_sheet(
                "RATES", suffix, f"OPUS RATES{tag}",
                [r.model_dump() for r in row_set.rates], ref_wb,
                rates_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet,
            ))
        if want_exploded and row_set.rates_port_port:
            results.append(_compare_keyed_sheet(
                "RATES PORT-PORT", suffix, f"OPUS RATES{tag} PORT-PORT",
                [r.model_dump() for r in row_set.rates_port_port], ref_wb,
                rates_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet,
            ))
        if row_set.arbs:
            results.append(_compare_keyed_sheet(
                "ARBS", suffix, f"OPUS ARBS{tag}",
                [r.model_dump() for r in row_set.arbs], ref_wb,
                arbs_row_key, cols.ARBS_ROW_FIELDS, read_arbs_sheet,
            ))
        if row_set.cmdt_notes:
            results.append(_compare_block_sheet(
                "CMDT NOTE", suffix, f"OPUS CMDT NOTE{tag}",
                [r.model_dump() for r in row_set.cmdt_notes], ref_wb,
                cols.CMDT_NOTE_ROW_FIELDS, read_cmdt_note_sheet,
            ))
        if row_set.special_notes:
            results.append(_compare_block_sheet(
                "SPECIAL NOTE", suffix, f"OPUS SPECIAL NOTE{tag}",
                [r.model_dump() for r in row_set.special_notes], ref_wb,
                cols.SPECIAL_NOTE_ROW_FIELDS, read_special_note_sheet,
            ))
    return results


def _render_results(results: list[dict]) -> None:
    if not results:
        st.info("Nothing to compare - the parsed MRG produced no rows for the sheet type(s) selected.")
        return

    st.markdown("#### Comparison summary")
    st.dataframe(
        [
            {
                "Sheet": r["sheet_type"],
                "Sub-lane": r["sub_lane"],
                "In reference?": "Yes" if r["found_in_reference"] else "No - sheet not found",
                "Matched": r["matched"] if r["matched"] is not None else "-",
                "Missing": len(r["missing"]),
                "Extra": len(r["extra"]),
                "Field mismatches": len(r["field_mismatches"]),
            }
            for r in results
        ],
        hide_index=True,
        width="stretch",
    )

    st.markdown("#### Details")
    for r in results:
        label = f"{r['sheet_type']} — {r['sub_lane']} ({r['sheet_name']})"
        with st.expander(label):
            if not r["found_in_reference"]:
                st.warning(
                    f"Reference workbook has no sheet matching **{r['sheet_name']}** - "
                    "every generated row is listed as extra."
                )
            if r["missing"]:
                st.markdown(f"**Missing** ({len(r['missing'])}, in reference but not generated)")
                st.dataframe(r["missing"][:50], hide_index=True, width="stretch")
            if r["extra"]:
                st.markdown(f"**Extra** ({len(r['extra'])}, generated but not in reference)")
                st.dataframe(r["extra"][:50], hide_index=True, width="stretch")
            if r["field_mismatches"]:
                st.markdown(f"**Field mismatches** ({len(r['field_mismatches'])})")
                st.dataframe(r["field_mismatches"][:50], hide_index=True, width="stretch")
            if not r["missing"] and not r["extra"] and not r["field_mismatches"]:
                st.success("No differences found.")


def render() -> None:
    state = _get_state()
    st.subheader("Compare: MRG vs. reference OPUS file")
    st.caption(
        "Upload the raw MRG rate sheet(s) and an existing OPUS-format Excel file "
        "(e.g. a filing someone already produced) to see where they diverge."
    )

    mrg_files = st.file_uploader(
        "Raw MRG rate sheet(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="compare_mrg_upload"
    )
    reference_file = st.file_uploader(
        "Reference OPUS-format Excel file (.xlsx)", type=["xlsx"], key="compare_reference_upload"
    )

    if not mrg_files or reference_file is None:
        st.info("Upload both the MRG file(s) and a reference OPUS file to continue.")
        return

    mrg_names = [f.name for f in mrg_files]
    mrg_payloads = [f.getvalue() for f in mrg_files]
    reference_payload = reference_file.getvalue()

    upload_key = fingerprint_uploads(mrg_names + [reference_file.name], mrg_payloads + [reference_payload])
    if upload_key != state.upload_key:
        state.upload_key = upload_key
        state.workbook = None
        state.classification_results = []
        state.selected_lane_id = None
        state.reference_workbook = None
        state.row_sets = None
        state.compare_results = None

    if state.workbook is None:
        try:
            state.workbook, state.classification_results = load_and_classify(mrg_payloads)
        except DuplicateSheetError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surfaced directly to the user, not swallowed
            st.error(f"Couldn't open one of the MRG files: {exc}")
            return

    if state.reference_workbook is None:
        try:
            state.reference_workbook = openpyxl.load_workbook(io.BytesIO(reference_payload), data_only=True)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Couldn't open the reference OPUS file: {exc}")
            return

    results = state.classification_results
    best = results[0] if results else None
    if best is None:
        st.error("No lane parsers are registered - nothing to classify against.")
        return

    lane_ids = [r.profile.lane_id for r in results]
    default_lane = state.selected_lane_id or best.profile.lane_id
    selected = st.selectbox(
        "Lane", options=lane_ids, index=lane_ids.index(default_lane) if default_lane in lane_ids else 0,
        help="Auto-selected from the best classification match; override if it's wrong.",
    )
    state.selected_lane_id = selected

    state.rates_mode = st.radio(
        "Generate MRG as:", options=RATES_MODE_OPTIONS,
        index=RATES_MODE_OPTIONS.index(state.rates_mode),
        help="Controls which of the two derived RATES forms gets compared - both come from the same parse.",
    )

    if st.button("Run Comparison", type="primary"):
        parser_cls = get_profile(state.selected_lane_id).parser_cls
        parser = parser_cls()
        with st.spinner("Parsing MRG..."):
            state.row_sets = run_parser(parser, state.workbook, MappingProfile())
        state.compare_results = _run_comparison(state.row_sets, state.reference_workbook, state.rates_mode)

    if state.compare_results is not None:
        _render_results(state.compare_results)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `./.venv/Scripts/python.exe -m pytest tests/test_compare_page.py -v`
Expected: PASS, 4 tests.

- [ ] **Step 5: Wire the mode toggle into `app.py`**

Replace the entire content of `mrg2opus/ui/app.py` with:

```python
"""Phase 2 Streamlit wizard entrypoint.

    ./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py

Two modes, selected at the top: "Convert" (the 4-step wizard:
upload+classify -> preview -> customize -> export) and "Compare"
(standalone: upload an MRG plus a reference OPUS file, see where they
diverge - see docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md).
"""
from __future__ import annotations

import streamlit as st

# Importing the lane modules registers their LayoutProfile as a side effect -
# same requirement as cli.py.
from mrg2opus.parsers import cse, eaf, laec, lawc, saf  # noqa: F401
from mrg2opus.ui import compare_page
from mrg2opus.ui.state import get_state
from mrg2opus.ui.steps import step1_upload, step2_preview, step3_customize, step4_export

STEP_LABELS = {
    1: "1. Upload & Classify",
    2: "2. Preview",
    3: "3. Customize",
    4: "4. Export",
}

STEP_RENDERERS = {
    1: step1_upload.render,
    2: step2_preview.render,
    3: step3_customize.render,
    4: step4_export.render,
}


def main() -> None:
    st.set_page_config(page_title="mrg2opus", layout="wide")
    st.title("MRG → OPUS Converter")

    mode = st.radio("Mode", options=["Convert", "Compare"], horizontal=True)
    st.divider()

    if mode == "Compare":
        compare_page.render()
        return

    state = get_state()

    cols = st.columns(len(STEP_LABELS))
    for col, (step_num, label) in zip(cols, STEP_LABELS.items()):
        with col:
            if step_num == state.step:
                st.markdown(f"**➤ {label}**")
            elif step_num < state.step:
                st.markdown(f"✅ {label}")
            else:
                st.markdown(f"{label}")
    st.divider()

    STEP_RENDERERS[state.step](state)


if __name__ == "__main__":
    main()
```

- [ ] **Step 6: Manually verify the Compare flow end-to-end in the browser**

Start the dev server (via the project's `.claude/launch.json` "streamlit-ui" config through the Browser preview tool). Switch Mode to "Compare". Upload `Sample MRGs with OPUS FORMATS/SAF.xlsx` as the MRG file, and the same file again as the reference OPUS file (it contains both the raw `SAF` sheet and the ground-truth `OPUS RATES`/`OPUS RATES PORT-PORT`/`OPUS CMDT NOTE` sheets, so this is a same-file self-comparison — the fastest way to confirm the whole flow works without needing a second real file). Confirm: lane auto-selects SAF, click "Run Comparison", and the summary table shows RATES and RATES PORT-PORT rows with `Matched` counts and low/zero `Missing`/`Extra` (this specific sample won't be a perfect 100% match on CMDT NOTE due to the same documented per-lane gaps the golden tests already accept — that's expected, not a bug in this feature). Then switch Mode back to "Convert" and confirm Step 1 still renders normally (the toggle doesn't clobber the wizard's own state).

- [ ] **Step 7: Run the full test suite**

Run: `./.venv/Scripts/python.exe -m pytest tests/ -v`
Expected: PASS, all tests from every prior task plus this task's 4 new ones.

- [ ] **Step 8: Commit**

```bash
git add mrg2opus/ui/compare_page.py mrg2opus/ui/app.py tests/test_compare_page.py
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Add standalone Compare mode: upload MRG + reference OPUS, see the diff"
```

---

## Task 7: Documentation

**Files:**
- Modify: `MIGRATION_NOTES.md`
- Modify: `README.md`

**Interfaces:** None — documentation only, no code.

- [ ] **Step 1: Add a MIGRATION_NOTES.md section**

Find the line `## 4. Shared building blocks (in \`parsers/common/\`)` in `MIGRATION_NOTES.md` and insert this new section immediately before it:

```markdown
## 3.11 Standalone Compare mode: MRG vs. reference OPUS file

User-requested, deliberately separate from the wizard (not a step in it,
and separate from the still-pending Phase 3 Audit Gate spec's internal
self-check rules - see that spec's own section once it's built): upload
raw MRG file(s) plus an existing OPUS-format Excel file, and see where
mrg2opus's own parse of the MRG diverges from that reference file, across
all 5 OPUS sheet types.

Promotes `tests/golden.py`'s sheet-reading and keyed-diff logic
(`read_rates_sheet`/`read_arbs_sheet`/etc., `diff_rates`/`rates_row_key`)
- previously test-only - into a new production module
`mrg2opus/audit/compare.py`, generalized as `diff_by_key(generated,
expected, key_fn, fields, ignore_fields)`. `golden.py` now delegates to
it (thin path-loading wrappers preserving every existing test file's
exact call signature - verified via grep before the refactor that
`_normalize` and `_normalize_cmdt_value` are both still directly imported
by different lane test files and are NOT touched by this refactor, only
the sheet readers and `diff_rates`/`rates_row_key` are).

CMDT NOTE / SPECIAL NOTE can't reuse the keyed diff - no reliable per-row
key exists (child rows share their parent's blank `header_seq`/
`note_seq`), and a reference file's row order isn't guaranteed to match
the generator's the way it is when golden tests always compare against
one exact known sample. `diff_cmdt_blocks()` reconstructs each side's
block structure independently (a non-blank `header_seq`/`note_seq` starts
a block; blank-seq rows belong to it - mirrors the writer's own fill-down
convention), keys blocks by the parent's `contents` text, and only
compares child rows positionally *within* a matched block - scoping the
fragile ordering assumption to one block instead of the whole sheet.

UI: a `st.radio("Mode", ["Convert", "Compare"])` at the top of `app.py`
(unrelated to and doesn't renumber the wizard's own 4 steps). Compare's
upload+classify logic is shared with the wizard's Step 1 via a new
`mrg2opus/ui/mrg_upload.py` (`fingerprint_uploads`, `load_and_classify`)
instead of two copies. A "Generate MRG as: Grouped (RATES) / Exploded
(RATES PORT-PORT) / Both" selector controls which of the two derived
RATES forms get compared - not a second parse path, since
`RatesPortPortRow` is always a deterministic transform of `RatesRow`
(`explode_rates_row()`). `CompareState` (in `compare_page.py`, not
`ui/state.py` - it has exactly one consumer) caches against a sha256
fingerprint of the MRG file(s) AND the reference file together, same
pattern as `WizardState.upload_key` (§3.10) - re-uploading either with
edited contents correctly invalidates the cached parse/comparison.

Full spec: `docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md`.
Test coverage: `tests/test_audit_compare.py` (readers, keyed diff, block
diff), `tests/test_compare_engine_regression.py` (the new engine against
all 5 bundled samples' own ground truth), `tests/test_mrg_upload.py`,
`tests/test_compare_page.py`.
```

- [ ] **Step 2: Update README.md**

Find the section describing the wizard steps in `README.md` (look for the paragraph starting "Opens a 4-step wizard") and add a new paragraph immediately after the existing "**Withdrawn locations are dropped...**" paragraph (search for that exact heading text) with this content:

```markdown
**Compare mode: check an existing OPUS file against its source MRG.** A
second top-level mode (switch at the top of the app, separate from the
4-step wizard) - upload the raw MRG file(s) plus an existing OPUS-format
Excel file (e.g. a filing someone already produced), and see a row-level
and field-level diff across all 5 OPUS sheet types the lane produces.
Choose whether to check the MRG as grouped (RATES), exploded (RATES
PORT-PORT), or both - both are always derived from the same parse, this
only controls what gets compared. See
`mrg2opus/audit/compare.py`/`mrg2opus/ui/compare_page.py`.
```

- [ ] **Step 3: Commit**

```bash
git add MIGRATION_NOTES.md README.md
git -c user.name="Ralf Barbuena" -c user.email="ralfjhonbarbuena@gmail.com" commit -m "Document the Compare mode in MIGRATION_NOTES.md and README.md"
```

---

## Final verification

After Task 7's commit, run the complete suite once more to confirm the whole plan leaves the project in a fully green state:

```bash
./.venv/Scripts/python.exe -m pytest tests/ -v
```

Expected: every test passes — the pre-plan 54, plus 7 (Task 1) + 4 (Task 2) + 5 (Task 4) + 5 (Task 5) + 4 (Task 6) = 79 total.
