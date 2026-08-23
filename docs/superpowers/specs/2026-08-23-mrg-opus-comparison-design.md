# MRG-vs-OPUS Comparison Procedure — Design

Date: 2026-08-23
Status: Approved, ready for implementation plan
Companion spec: `2026-08-23-phase3-audit-gate-design.md` (unchanged, still
approved, still to be built — this is additive, not a replacement; see
Context below)

## Context

While brainstorming the Phase 3 Audit Gate spec (internal self-check
rules run inside the wizard, blocking export on serious findings), the
user asked for a second, independent capability: instead of only
checking a freshly-generated `OpusRowSet` against structural rules, let a
user upload a raw MRG **and** an existing OPUS-format Excel file (e.g. a
filing someone else already produced, or a prior run's output) and see
where the two diverge. This is explicitly a separate procedure, not a
wizard step — the user's own words: "instead of being a phase, be a
separate procedure."

Two decisions from brainstorming fixed the scope:
- **Keep both** — the audit-gate spec's 4 internal rules still get built
  as designed; this comparison procedure is additional, not a
  replacement.
- **All 5 OPUS sheet types** get compared in v1 (RATES, RATES PORT-PORT,
  ARBS, CMDT NOTE, SPECIAL NOTE) — not a RATES-only first pass.

This reuses, rather than reinvents, comparison logic that already exists
and is already proven: `tests/golden.py`'s sheet readers and
`diff_rates()` are the exact mechanism every lane's golden tests already
use to verify parser output against the 5 bundled samples' own
ground-truth OPUS sheets. This spec promotes that logic from test-only
code into a production module the UI (and the tests) both call.

## Goals

- Let a user upload one or more raw MRG files plus one reference
  OPUS-format Excel file, and see a row-level and field-level diff
  between what mrg2opus would generate from the MRG and what the
  reference file actually contains — across all 5 OPUS sheet types the
  lane produces.
- Reuse the proven `rates_row_key`/`diff_rates` matching logic for
  RATES, RATES PORT-PORT, and (via a new `arbs_row_key`) ARBS, rather
  than inventing new comparison semantics for shapes that already have
  a tested one.
- Handle CMDT NOTE / SPECIAL NOTE correctly despite them having no
  natural per-row key and depending on fill-down order — via a
  block-reconstruction approach, not a fragile whole-sheet positional
  diff (which only golden tests can safely assume, since they always
  compare against one exact known sample).
- Consolidate `tests/golden.py`'s reader/diff logic to import from the
  new production module, so there is one implementation instead of two
  that can silently drift apart.
- Let the user choose, before running a comparison, whether they want
  the MRG evaluated as grouped (RATES), exploded (RATES PORT-PORT), or
  both — framed as a generation choice even though, mechanically, both
  forms are always derived from the same parse (see Non-goals).

## Non-goals

- **Not a second generation path.** `RatesPortPortRow` is always a
  deterministic transform of `RatesRow` (`explode_rates_row()`) — there
  is no way for "exploded-only" to produce different *data* than
  "grouped, then explode." The Grouped/Exploded/Both selector controls
  which derived form(s) get compared, not a different parse.
- **No CLI subcommand.** UI-only per the brainstorming decision.
- **Does not replace or modify the Audit Gate spec's wizard-integrated
  rules.** Those are unaffected; this is a fully separate mode.
- **No persistent comparison history.** Each run is ephemeral, shown
  once in the session, same as the rest of this app's current scope.

## Architecture

### `mrg2opus/audit/compare.py` (new)

**Sheet readers**, promoted from `tests/golden.py` with one signature
change: they take an already-loaded `Workbook`, not a file `Path` — the
caller loads the workbook (from a test fixture path, or from a
Streamlit-uploaded file's bytes via `io.BytesIO`), so the same reader
serves both without duplicating load logic.

```python
def find_sheet(wb: Workbook, sheet_name: str) -> str: ...       # was _find_sheet
def read_rates_sheet(wb: Workbook, sheet_name: str) -> list[dict]: ...
def read_arbs_sheet(wb: Workbook, sheet_name: str = cols.SHEET_NAME_ARBS) -> list[dict]: ...
def read_cmdt_note_sheet(wb: Workbook, sheet_name: str) -> list[dict]: ...
def read_special_note_sheet(wb: Workbook, sheet_name: str = cols.SHEET_NAME_SPECIAL_NOTE) -> list[dict]: ...
```

**Keyed diff**, generalized from `diff_rates()`:

```python
@dataclass
class KeyedDiffResult:
    matched: int
    missing: set[tuple]          # keys present in reference, not generated
    extra: set[tuple]            # keys present in generated, not reference
    field_mismatches: list[tuple[tuple, str, Any, Any]]  # (key, field, generated_value, reference_value)

def diff_by_key(generated: list[dict], expected: list[dict], key_fn, fields: list[str],
                 ignore_fields: set[str] = frozenset()) -> KeyedDiffResult: ...

def rates_row_key(row: dict) -> tuple: ...   # unchanged, moved from golden.py
def arbs_row_key(row: dict) -> tuple:        # new - promoted from test_parsers_cse.py's inline _arbs_key
    return (row.get("point"), row.get("over"), row.get("per"))
```

RATES and RATES PORT-PORT both call `diff_by_key(..., key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS)`.
ARBS calls `diff_by_key(..., key_fn=arbs_row_key, fields=cols.ARBS_ROW_FIELDS)`.

### CMDT NOTE / SPECIAL NOTE: block-based diff

Neither sheet has a reliable per-row key (child rows share their
parent's blank `header_seq`/`note_seq`), and a reference file's row
order isn't guaranteed to match the generator's — unlike golden tests,
which only ever compare against one exact known sample and can safely
assume matching order.

```python
@dataclass
class CmdtBlock:
    key: str                # the parent row's `contents` text, stripped
    parent: dict
    children: list[dict]

def reconstruct_blocks(rows: list[dict]) -> list[CmdtBlock]:
    """A row with non-None header_seq OR note_seq starts a new block
    (mirrors the writer's own fill-down convention); every following row
    until the next such marker belongs to it."""

@dataclass
class BlockDiffResult:
    missing_blocks: list[str]     # contents-text keys in reference, not generated
    extra_blocks: list[str]       # contents-text keys in generated, not reference
    field_mismatches: list[tuple[str, int, str, Any, Any]]  # (block_key, child_index, field, gen, ref)

def diff_cmdt_blocks(generated: list[dict], expected: list[dict],
                      fields: list[str]) -> BlockDiffResult:
    """Blocks matched by contents-text key; missing/extra reported at
    block granularity. For matched blocks, child rows compared
    positionally (safe here - the fragile assumption is scoped to one
    block's internal order, not the whole sheet)."""
```

CMDT NOTE uses `cols.CMDT_NOTE_ROW_FIELDS`; SPECIAL NOTE uses
`cols.SPECIAL_NOTE_ROW_FIELDS` (same function, different field list).

### `tests/golden.py` refactor

`golden.py`'s reader functions and `diff_rates()`/`rates_row_key()`
become thin re-exports from `audit/compare.py` (or the module is edited
to import and call the production versions directly). Existing test call
sites load the workbook once (`openpyxl.load_workbook(path, ...)`) and
pass it to the readers instead of passing a path — a small, mechanical
edit across the golden-test files, behavior unchanged. This is in scope
because it's the literal point of the feature (promoting proven test
logic to production), not an unrelated refactor — but it touches every
lane's passing golden tests, so it gets its own verification pass in the
implementation plan before anything else builds on top of it.

## UI wiring

- `app.py` gains a top-level mode selector: `st.radio("Mode", ["Convert", "Compare"])`,
  stored outside `WizardState` (Compare doesn't participate in the
  5-step wizard at all). Selecting "Compare" renders a new
  `ui/compare_page.py` instead of `STEP_RENDERERS[state.step]`.
- Shared upload+classify logic (merge multiple MRG files, run
  `classify_all`, let the user confirm/override the lane) is factored
  out of `step1_upload.py` into `ui/mrg_upload.py` so both the wizard and
  Compare call one implementation instead of two copies drifting apart.
- Compare flow, top to bottom:
  1. Upload MRG file(s) — via the shared helper.
  2. Upload one reference OPUS-format Excel file.
  3. **"Generate MRG as:"** radio — Grouped (RATES) / Exploded (RATES
     PORT-PORT) / Both, default Both. Controls which of the two derived
     forms gets compared (see Non-goals — not a second parse path).
  4. "Run Comparison" button.
  5. Results: a summary table (sheet type, sub-lane, matched/missing/extra
     counts), each row expandable to the actual missing/extra rows and
     field-level mismatches. A sheet type is skipped entirely (not shown
     as a zero-row false negative) when neither side has any data for
     it — e.g. a lane with no ARBS.
- New `CompareState` dataclass, defined directly in `compare_page.py`
  (not `ui/state.py` — unlike `WizardState`, which every step module
  imports, this has exactly one consumer, so it doesn't belong in the
  shared state module). Same shape/spirit as `WizardState`: caches the
  parsed MRG `OpusRowSet` and the loaded reference `Workbook` keyed
  against a **sha256 content fingerprint of both uploads together** — the
  same pattern `WizardState.upload_key` uses (§3.10 of
  MIGRATION_NOTES.md), so this new feature doesn't reintroduce the exact
  stale-cache bug just fixed in the main wizard. Held in
  `st.session_state.compare` (parallel to `get_state()`'s
  `st.session_state.wizard`), not inside `WizardState` itself.

## Testing

- `tests/test_audit_compare.py` (new):
  - `diff_by_key`: missing-key, extra-key, and field-mismatch cases with
    small synthetic dicts.
  - `diff_cmdt_blocks`: matched blocks (identical), a missing block, an
    extra block, and a positional field mismatch within one matched
    block.
- One regression test per bundled sample (SAF, EAF, CSE, LAEC, LAWC):
  parse the sample's raw sheet(s), read the SAME workbook's own
  ground-truth OPUS sheets as the "reference," run the new compare
  engine end-to-end for every sheet type the lane produces, and assert
  the match rate lines up with what the existing golden tests already
  document (including each lane's already-known, already-documented
  gaps — reuse the same `ignore_fields`/known-gap exclusions those tests
  use, don't re-litigate them here). This both verifies the new engine
  and re-derives the existing golden tests through the new production
  code path as a bonus consistency check.

## Files touched

- New: `mrg2opus/audit/compare.py`, `mrg2opus/ui/compare_page.py`,
  `mrg2opus/ui/mrg_upload.py`, `tests/test_audit_compare.py`.
- Edited: `mrg2opus/ui/app.py` (mode selector, routing),
  `mrg2opus/ui/steps/step1_upload.py` (delegates to the new shared
  helper), `tests/golden.py` (imports from `compare.py` instead of
  reimplementing).
- Documentation: `MIGRATION_NOTES.md` (new section), `README.md`
  (Compare mode description).
