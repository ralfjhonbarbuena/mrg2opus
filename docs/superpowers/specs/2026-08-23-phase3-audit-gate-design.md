# Phase 3 Audit Gate — Design

Date: 2026-08-23
Status: Approved, ready for implementation plan

## Context

mrg2opus converts raw carrier MRG rate sheets into structured OPUS filing
Excel files, for billing-relevant freight data where silent
misclassification is costly. Phase 1 (parsing engine, 5 lanes) and Phase 2
(Streamlit wizard: Upload → Preview → Customize → Export) are done. The
original project plan called for a Phase 3 "Audit Gate": a review step
between Customize and Export that surfaces integrity problems in the
parsed output before the user downloads it, so obviously-wrong data
doesn't ship silently.

`mrg2opus/audit/__init__.py` exists today only as an empty package stub —
no rule functions, no UI slot. This spec covers building the first,
deliberately small set of rules and wiring them into the wizard.

This work does not depend on the user's promised follow-up MRG/OPUS
examples (WAF, West Asia WAF, TAD Filing, LAEC's route note/ARBS/SPECIAL
NOTE gaps) — it operates purely on the existing 5 implemented lanes'
output shape, which is already fully specified.

## Goals

- Catch data-integrity problems that are structurally detectable from
  `OpusRowSet` (or, for one rule, the raw uploaded workbook) without
  needing any new ground-truth samples.
- Block export on the two failure modes serious enough to warrant it
  (missing ARBS for an ARBS-eligible lane; a rate row where every
  container size is unfiled), with an explicit, visible
  acknowledge-and-proceed override — never a silent block with no way
  forward.
- Surface two more patterns informationally (conflicting duplicate
  routes; freetime present in source but never extracted) without
  blocking, since neither is unambiguously wrong.
- Keep each rule an independently readable, independently testable pure
  function — this is a first pass of 4 rules, but the shape needs to
  scale to a 5th/6th without becoming a monolith.

## Non-goals (explicitly out of scope for this pass)

- **Unmapped POL/POD detection.** Today an unresolvable raw location
  silently drops the row entirely rather than producing a flagged one —
  there's nothing to check yet. Needs Pending Task #1 (Location Bank
  match-confidence tracking, a parser contract change across all 5
  lanes) first. Not attempted here.
- **Commodity group code validation against a bank.** The commodity bank
  was removed when codes became fully user-owned (§3.6 of
  MIGRATION_NOTES.md) — there's no registry left to validate against.
- **Persistent audit-trail logging.** The original plan sketch mentioned
  a "logged acknowledge-and-export-anyway override." No logging
  infrastructure exists anywhere in this app (no server, no audit log
  file) and building one is a separate concern. The acknowledgment is
  made visible on the Export screen instead (see UI section) — honest
  in-session visibility, not a persistent record.
- **New lanes, new rules beyond the 4 listed here.** Deliberately small
  first pass; more rules are a cheap follow-up once these prove out.

## Architecture

### `Finding` model (new, `mrg2opus/audit/checks.py`)

```python
from typing import Literal
from pydantic import BaseModel

class Finding(BaseModel):
    rule_id: str
    severity: Literal["blocking", "info"]
    sub_lane: str            # row_sets suffix; "" for single-sublane lanes (e.g. CSE)
    summary: str              # one-line, shown in the findings table
    detail_rows: list[dict] = []  # small dicts identifying affected rows, for an expander
```

### Rule functions

Each rule is a plain function, not a class — with only 4 rules, a
`Rule` object with metadata (id/severity/description as fields, `.check()`
as a method) is pure ceremony with no current benefit. A single
monolithic `audit()` function was also considered and rejected: it fights
the same isolation principle the rest of `parsers/common/` already
follows (small, single-purpose, independently testable units), and
doesn't scale cleanly to a 5th/6th rule later.

```python
def check_missing_arbs(lane_id: str, row_sets: dict[str, OpusRowSet]) -> list[Finding]: ...
def check_zero_dollar_rates(row_sets: dict[str, OpusRowSet]) -> list[Finding]: ...
def check_duplicate_routes(row_sets: dict[str, OpusRowSet]) -> list[Finding]: ...
def check_freetime_not_extracted(workbook: Workbook) -> list[Finding]: ...

ALL_CHECKS = [check_missing_arbs, check_zero_dollar_rates, check_duplicate_routes, check_freetime_not_extracted]

def run_audit(lane_id: str, row_sets: dict[str, OpusRowSet], workbook: Workbook) -> list[Finding]:
    """Runs every check, returns the combined finding list."""
```

`check_freetime_not_extracted` needs the raw workbook, not just
`OpusRowSet` — nothing in the OPUS schema carries freetime data at all
(project-wide known gap, §6 of MIGRATION_NOTES.md), so the only way to
detect "freetime info exists in source" is to scan the raw sheets
directly. This is why `run_audit`'s signature takes the workbook
alongside the parsed rows, not `OpusRowSet` alone as the original plan
sketch implied.

### Rule logic

- **`check_missing_arbs`**: `ARBS_ELIGIBLE_LANES = {"CSE", "LAEC"}`,
  hardcoded with a comment (matches the project's existing convention of
  documenting verified-but-not-derivable constants inline, e.g.
  `cse.py`'s `CMDT_NOTE_CHARGE_CODES`). For each sub-lane, if
  `lane_id in ARBS_ELIGIBLE_LANES` and `row_set.arbs` is empty → one
  blocking `Finding` (`rule_id="missing_arbs"`).
- **`check_zero_dollar_rates`**: a `RatesRow` where `rate_20`, `rate_40`,
  `rate_40hc`, AND `rate_45` are all `None` or `0` — not just one missing
  size, which is normal (a lane may legitimately never offer 45'). One
  blocking `Finding` per sub-lane (`rule_id="zero_dollar_rate"`) grouping
  every offending row, `detail_rows` capped (first 20) with the total
  count stated in `summary` so a systemic break doesn't flood the UI.
- **`check_duplicate_routes`**: within each sub-lane's `rates` list
  (never across sub-lanes — EAF's TZDAR/KEMBA are separate filings),
  group by `(commodity_group_code, origin_code, destination_code,
  prefix, cgo_type)`. Flag only groups where the grouped rows' rate
  values actually differ — identical duplicates aren't a conflict, don't
  flag those. Informational (`rule_id="duplicate_route"`).
- **`check_freetime_not_extracted`**: scan every non-`OPUS`-prefixed
  sheet (same skip pattern as `registry.py`'s `_title_keyword_score`) for
  cells containing "free time", "freetime", "demurrage", or "detention"
  (case-insensitive substring). One informational `Finding`
  (`rule_id="freetime_not_extracted"`) on the first hit, naming the sheet
  and cell reference — not exhaustive, just a pointer.

## UI wiring

Wizard goes from 4 steps to 5: **Upload → Preview → Customize → Audit
(new) → Export**.

- `mrg2opus/ui/steps/step4_export.py` → renamed to `step5_export.py`.
- New `mrg2opus/ui/steps/step4_audit.py`.
- `app.py`'s `STEP_LABELS`/`STEP_RENDERERS` become 5 entries; imports
  updated for the renamed export module.
- `step3_customize.py`'s "Apply & Continue" already sets `state.step = 4`
  — that lands on the new Audit step with no change needed there.
  `step4_audit.py`'s own "Continue to Export →" sets `state.step = 5`.

`WizardState` gains `audit_acknowledged: bool = False`. Reset to `False`
everywhere `row_sets` is invalidated or reassigned — Step 1's
upload-changed reset block, Step 2's "↻ Re-parse from source" button, and
Step 3's "Apply & Continue" re-run. A stale acknowledgment must never
carry over to a genuinely different parse.

`step4_audit.py` recomputes findings fresh on every render
(`run_audit(state.selected_lane_id, state.row_sets, state.workbook)`) —
cheap, pure functions over already-parsed in-memory data — rather than
caching the findings list itself. Deliberately avoiding a second
cache-invalidation surface, having just fixed one this session (§3.10 of
MIGRATION_NOTES.md: the wizard's row_sets cache going stale).

Findings render as a table (severity, sub-lane, rule, summary), each row
expandable to show its `detail_rows` preview. If any `severity="blocking"`
finding exists, a checkbox ("I've reviewed these and want to export
anyway") must be checked before "Continue to Export →" is enabled. If the
user proceeds with blocking findings acknowledged, Step 5 (Export) shows
a visible notice stating how many blocking findings were present and
acknowledged — the closest honest equivalent to "logged" this app's
architecture supports without adding persistent logging infrastructure.

## Testing

`tests/test_audit_checks.py` — per rule, a positive case (produces the
expected `Finding`) and a clean case (no false positive), built from
small hand-crafted `OpusRowSet`/`Workbook` fixtures, matching this
session's CSE regression-test style (synthetic in-memory workbooks, not
golden-diff against samples — these rules are about detecting anomalies,
not reproducing ground truth).

One additional regression test: run `run_audit()` against each of the 5
bundled samples' actual parsed output (SAF, EAF, CSE, LAEC, LAWC) and
assert zero blocking findings. Since these are known-good ground truth,
this becomes a tripwire — if a future change ever silently breaks ARBS
generation for CSE/LAEC or introduces zero-dollar rows anywhere, this
test catches it immediately without needing a new fixture.

## Files touched

- New: `mrg2opus/audit/checks.py`, `mrg2opus/ui/steps/step4_audit.py`,
  `tests/test_audit_checks.py`.
- Renamed: `mrg2opus/ui/steps/step4_export.py` →
  `mrg2opus/ui/steps/step5_export.py`.
- Edited: `mrg2opus/ui/app.py` (step dict, imports),
  `mrg2opus/ui/state.py` (`audit_acknowledged` field),
  `mrg2opus/ui/steps/step1_upload.py`,
  `mrg2opus/ui/steps/step2_preview.py`,
  `mrg2opus/ui/steps/step3_customize.py` (reset points for
  `audit_acknowledged`).
- Documentation: `MIGRATION_NOTES.md` (new §3.11 or next available
  number), `README.md` (wizard step count/description).
