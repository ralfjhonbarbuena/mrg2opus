# Migration Note — mrg2opus (MRG → OPUS Excel Converter)

Handoff snapshot for the next session. Read this first, then `README.md`
for setup/run commands, then the plan at
`C:\Users\romsae-desktop\.claude\plans\build-me-an-app-lexical-ocean.md`
for the original full design.

## 1. What this is

A Python app that converts raw carrier MRG (Master Rate Guideline) Excel
rate sheets into the structured OPUS filing format the company's rating
system consumes. Working directory:
`C:\Users\romsae-desktop\claude\PROCESS INNOVATION HACKATHON` (not a git
repo). Ground truth for everything built so far comes from the 5 paired
sample workbooks in `Sample MRGs with OPUS FORMATS/` — there is no other
spec; every schema field and business rule was reverse-engineered directly
from those files' own `OPUS *` sheets.

## 2. Architecture (unchanged from the plan, now implemented)

```
mrg2opus/
  schema/            # opus_rows.py (pydantic row models), opus_columns.py
                      # (canonical header text/order per sheet), charge_codes.py
  excel_io/           # style_utils.py (strikethrough/fill exclusion),
                      # writer.py (2-row RATES header, single-row NOTE headers,
                      # write_opus_workbook / write_opus_workbook_multi),
                      # merge.py (multi-file upload -> one workbook, see §3.7)
  parsers/
    base.py            # BaseMRGParser ABC: detect(), parse_raw(), to_opus_rows(),
                        # run(), run_multi() (for multi-sub-lane lanes like EAF)
    registry.py         # LayoutProfile registry + classify()/classify_all()/
                        # all_profiles()/get_profile() (Step 1 lane detection)
    common/              # shared building blocks (see §4)
    saf.py, eaf.py, cse.py, laec.py, lawc.py   # one module per lane, self-registering
  location_bank/       # SQLite store, rapidfuzz matcher, bootstrap scripts,
                        # known_aliases.py (curated spelling-variant fixes)
  presets/              # MappingProfile (Step 3 customization contract) + store.py
                        # (JSON-file-per-profile save/load/list/delete)
  ui/                    # Phase 2 Streamlit wizard - see §3.5
    app.py                 # entrypoint, session-state step router
    state.py                # WizardState dataclass + get_state()/reset_state()/goto()
    parsing.py               # run_parser() - run_multi() + apply commodity_group_order, §3.7
    commodity_utils.py        # distinct_commodity_groups(), shared by step2/step3
    steps/                     # one render(state) module per wizard step
  cli.py                 # python -m mrg2opus.cli parse <in.xlsx> --out <out.xlsx>
  config/
    container_maps/{lane}.yaml   # raw container-label → OPUS size-slot map
    group_codes/{lane}.yaml       # regional origin-group shorthand (FEBP, WPRD)
tests/                  # golden-file diffs against each lane's own ground truth
data/                   # generated/mutable (location_bank.sqlite3, presets/)
reference/               # curated (UN/LOCODE CSV goes here, not yet supplied)
```

Environment: Python 3.14.3 at
`C:\Users\romsae-desktop\AppData\Local\Python\bin\python.exe` (WindowsApps
`python`/`py` are broken store shims — don't use them). A `.venv` already
exists with all dependencies installed (openpyxl, pandas, pydantic, rapidfuzz,
streamlit, PyYAML, platformdirs, python-dateutil, pytest).

## 3. Per-lane status

| Lane | Status | Sub-structure | Tests |
|---|---|---|---|
| **SAF** | ✅ Done, exact match | Single raw sheet, 2 PODs (Durban/Cape Town) | 3/3 pass |
| **EAF** | ✅ Done, exact match | 2 sub-lanes (TZDAR/KEMBA), each a separate real-world file — see `run_multi()` | 8/8 pass |
| **CSE** | ✅ Done, near-exact (documented gaps) | 3 rate grids + 2 reefer sheets + in-gauge + Yangtze ARBS + DG surcharges→SPECIAL NOTE. Richest lane. Each raw sheet now defaults to its own commodity description (see §3.6). "CSE" grid sheet now tolerates a shifted header row + wider real-file destination/data extent (see §3.8). Withdrawn-location exclusion (see §3.9). | 9/9 pass |
| **LAEC** | ✅ Done, near-exact (documented gaps) | Same shape as CSE but doubled: every sheet splits into Non-ISC/ISC sections. "R5 NOR" now defaults to its own commodity description (see §3.6). | 4/4 pass |
| **LAWC** | ✅ Done, exact match | 4 Dry-shaped grids (main/SEA/ISC/OOG) + Reefer + NOR sheets, each its own commodity group. OOG has 4 equipment-type column-pairs mapping to Prefix O/F twins or F-only. RATES PORT-PORT uses a completely different commodity code namespace than RATES for the same data. Reefer/NOR now default to their own commodity descriptions (see §3.6). | 4/4 pass |
| **WAF** | ✅ Done, exact match | Single raw sheet, 9-POD x 3-container grid over ~41 fixed origins. RATES + CMDT NOTE only (no ARBS/SPECIAL NOTE/RN). D/DR rows each also file an identical D/DG duplicate under their own commodity group (see §3.19). | 6/6 pass |
| **AUEC** | ✅ Done, exact match (FAK only) | 3 raw sheets sharing one 48-origin grid, 2 commodity groups (main + NZJ). RATES + CMDT NOTE only. Adds RAD ("Reefer As Dry") container type, rail-routing transmode, PRDA/PRDB footnote-expanded regional groups, and per-charge-code POL scoping on CMDT NOTE children (see §3.21). TIER 1 variant not yet verified. | 7/7 pass |
| **AUWC** | ✅ Done, exact match (FAK only) | 2 raw sheets sharing one 49-origin grid (same PRDA/PRDB groups as AUEC). 3 commodity groups (main, RF, NOR each separate - unlike AUEC's folded-in RF/RAD). Same charge/duplicate-EFS list as AUEC but scoped on POR, not POL (see §3.22). TIER 1 variant not yet verified. | 8/8 pass |

Full test suite: **108/108 passing** (`./.venv/Scripts/python.exe -m pytest tests/ -v`, ~2 min) - see §3.15 for the 2026-08-26 rebuild onto `reference/` ground truth (SAF has no golden tests anymore; `test_compare_engine_regression.py`/`golden.py`/`conftest.py` were removed as redundant/dead), §3.16 for the `excluded_charge_codes` feature added the same day (`test_cmdt_notes.py` + 1 EAF end-to-end test), §3.17 for sequential default commodity codes (`test_commodity_utils.py`), §3.18 for the CSE 2-file merge fix (4 new `test_excel_io_merge.py` tests + CSE's own golden tests now cover the full 2-file merge), §3.19 for the new West Africa WAF lane (`test_parsers_waf.py`, 7 tests), §3.20 for the RFA effective/expiry date override (4 new `test_cmdt_notes.py` cases), §3.21 for the new AUS NEA to AUEC FAK lane (`test_parsers_auec.py`, 7 tests), and §3.22 for the new AUS NEA to AUWC FAK lane (`test_parsers_auwc.py`, 8 tests).

## 3.5 Phase 2 — Streamlit wizard UI

Built and verified end-to-end (both in-browser and via direct Python calls
against every sample lane, see below). Run it with:

```bash
./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py
```

A `.claude/launch.json` config (`streamlit-ui`) exists so it can be opened
via the browser-preview tool without retyping the command.

4-step linear wizard, session state held in `st.session_state.wizard`
(`ui/state.py::WizardState`), one render module per step under `ui/steps/`:

1. **Upload & Classify** (`step1_upload.py`) - `.xlsx` upload, runs
   `registry.classify_all()` (new - returns every lane's score, not just
   the winner) and shows the full ranked breakdown table. Auto-selects the
   best match but always shows a manual lane picker (forced into view when
   confidence is below the profile's threshold) - a wrong classification
   silently produces wrong freight rates, so this is never fully hands-off.
2. **Preview** (`step2_preview.py`) - runs `parser.run_multi()` with the
   default `MappingProfile`, shows row counts per sub-lane/sheet and a
   `st.dataframe` preview (capped at 500 rows) of whichever sheet the user
   picks.
3. **Customize** (`step3_customize.py`) - `st.data_editor` for commodity
   group code/description/CMDT-Seq overrides, seeded from
   `WizardState.default_commodity_groups` (a one-time snapshot of the
   parser's own default `(code, description)` pairs, captured right after
   Step 2's first parse - NOT re-derived from `state.row_sets` on every
   visit, since after an override is applied `row_sets` reflects the
   OVERRIDDEN codes and re-deriving from it would break a second round of
   edits; see `ui/commodity_utils.py`). There is no commodity code/
   description registry to read from either way - see §3.6. Also
   checkboxes to skip individual output sheets, and named-preset save/load
   via `presets/store.py`. Re-runs the parser with the edited
   `MappingProfile` before advancing.
4. **Export** (`step4_export.py`) - writes the (possibly sheet-filtered)
   result via `excel_io.writer.write_opus_workbook_multi()` to a temp
   file, reads it back as bytes, and offers `st.download_button`.

Verified via the in-app browser preview (full click-through with SAF.xlsx,
injected through the file input via a `DataTransfer`/`File` since this
sandboxed browser has no native file-picker) AND via direct Python calls
to the same step-module functions against EAF (2 sub-lanes) and CSE (ARBS
+ SPECIAL NOTE) to confirm the multi-sub-lane and richest-lane cases don't
just work for SAF's simple case. `skip_output_sheets` was verified to
actually remove the sheet from the written workbook, not just hide it from
a checkbox.

**Deliberately NOT built in this pass** (scope cuts, not oversights):
- **Location Bank low-confidence match review** (the plan's "top-5
  rapidfuzz candidates, confirm/override, write back as
  `source=manual_override`" flow). Building this properly needs every
  parser's `to_opus_rows()` to surface *which* raw tokens it resolved
  and at what confidence, not just silently drop rows it couldn't
  resolve - none of the 5 lane parsers currently return that
  information. Doing this well means touching every parser's return
  contract, which is a real design decision, not a UI-only addition -
  flagging for a deliberate follow-up rather than bolting on a
  half-working version.
- **Step 3.5 Audit Gate**: the findings/rules themselves are Phase 3
  scope (`audit/checks.py` is still just a package stub); the UI step
  slot for it can be added once there's something to show.
- Skipping individual RAW sheets isn't supported. `MappingProfile` used to
  carry a `skip_raw_sheets` field for this, but nothing ever read it - no
  parser honored it, and the UI never exposed it, so it was removed rather
  than left as a field that looks like a working control. Only
  `skip_output_sheets` exists, and it's wired into the writer path
  (`step4_export.py::_apply_skips()`). Re-adding raw-sheet skipping means
  threading it through each parser's sheet loop, not just restoring the
  field.

## 3.6 Commodity codes/descriptions are user-owned; every sheet defaults to
its own description; `type` is always "C"

Three changes made in the same later session as the Phase 2 UI, all direct
user directives, not derived from any sample. This section describes the
FINAL state after two rounds of iteration - the first round (code
overrides) shipped, then a second request ("every sheet should have its
own default description") forced a rework of the override key scheme
itself, described together with the original changes below rather than as
a separate stale entry.

**`commodity_bank/` package removed entirely** (it mined a JSON `{code:
description}` store from the 5 samples; nothing ever actually read from it
- the UI's commodity table was always built from parsed output, not this
file - and it had a known lane-namespacing bug, see the old §6 "Real bug"
entry, now moot). Commodity group codes/descriptions are meant to come
from the user, not a registry.

**Every raw sheet defaults to its own description now, not a description
shared with sibling sheets.** Three lanes had a raw sheet silently folded
into another sheet's combined description by default:
- LAWC: "Reefer" and "LAWC NOR" used to share the main dry grid's
  description ("China_TWN_SIN_HKG Dry and DG & REEFER & NOR", `G0001`).
- CSE: "NOR(PA)" and "CSE VE" used to share "CSE"'s description ("FAK &
  DG & NOR", `G0001`); "NOR (MAOVLD)" used to share "CSE (MAOVLD)"'s
  ("FAK - MAOVLD & MAOVLD NOR", `G0002`).
- LAEC: "R5 NOR" used to share the Non-ISC main group's description ("FAK
  & DG & NOR_NON-ISC", `G0015`).

Each of these now defaults to its OWN description (its own raw sheet
name), and the "parent" group's own description was trimmed to no longer
reference the sub-sheet that used to be folded in (e.g. LAEC's Non-ISC
main went from "FAK & DG & NOR_NON-ISC" to "FAK & DG_NON-ISC"). Each split
sub-group still defaults to the SAME `commodity_group_code` as its parent
(only description was asked to split, not code) and reuses the parent's
charge codes/validity for its own `OPUS CMDT NOTE` block (no independent
ground truth exists to derive a different set from).

**Merging is symmetric and general, not just "undo the split":** if the
user overrides ANY two groups within a lane (not just ones that used to be
bundled) to the exact same description, they collapse into ONE `OPUS CMDT
NOTE` block, with charge codes unioned (deduped, order-preserving) - see
`parsers/common/commodity.py::CommodityNoteSpec` /
`merge_note_specs()` / `build_notes_by_description()`. Verified as an exact
regression check per lane: overriding every split sub-group's description
back to the ORIGINAL combined ground-truth text reconstructs that lane's
`OPUS CMDT NOTE` sheet byte-for-byte (`test_*_cmdt_note_merges_when_descriptions_match`
in each of `test_parsers_{cse,laec,lawc}.py`). Implementation-wise this
required a real restructuring of each affected parser's `to_opus_rows()`:
notes used to be built inline as each commodity group's rows were built
(so `row.commodity_note` could be stamped immediately); now every group's
`CommodityNoteSpec` (description, validity, charge codes) is collected
FIRST across the whole lane, `build_notes_by_description()` runs ONCE at
the end (merging same-description specs before building), and a final
pass over the already-built `rates` list stamps `row.commodity_note` by
each row's own (already-final) `commodity_group_description`. Two lane-
specific wrinkles: CSE's own ground truth never populates
`commodity_note` on RATES rows at all (verified - leaves it blank
everywhere), so CSE's `to_opus_rows()` builds notes for the `OPUS CMDT
NOTE` sheet but skips the row-stamping pass entirely; LAWC's "CSE VE"-
equivalent split (multiple raw sheets feeding one `data.grid_rows` dict
key) required widening that dict's key from a 3-tuple to a 4-tuple, later
simplified to `(default_description, cmdt_seq, default_code)` once
description became the universal lookup key (see below) - dropping the
separate synthetic key that had briefly existed (e.g. `"G0001::VE"`,
`"G0001::REEFER"`) in favor of just using the description itself.

**All three `MappingProfile.commodity_*_overrides` dicts (code,
description, cmdt_seq) are now keyed by a group's DEFAULT description, not
its code.** This was a REVISION of the first-round design (which keyed by
code) - a real bug was caught only after wiring the UI: `ui/commodity_utils.py::
distinct_commodity_groups()` deduped by `commodity_group_code`, so any two
sub-groups sharing a default code (e.g. LAWC's main dry grid/"Reefer"/
"LAWC NOR", all defaulting to `G0001`) silently collapsed into ONE row in
the Step 3 editor - the user could never see or edit the hidden ones. The
"Apply" step's override-dict construction had the identical bug (also
keyed by code), so even fixing the display alone wouldn't have been
enough. Description is the right key because it's the one identity
ALWAYS unique per group by construction (every raw sheet gets its own
default description now) and directly visible to the UI - no synthetic
key needed. `parsers/common/commodity.py::resolve_commodity_code()` /
`resolve_commodity_description()` take the default description as the
lookup key (code resolution keeps a separate `default_code` parameter
since a group's default CODE can still legitimately be shared with
others, unlike the lookup key). Every one of the 5 lane parsers' override
call sites was updated to pass the default description, not code -
mechanical but not uniform (CSE/LAWC needed to compute
`variant_output_code`/`variant_cmdt_seq` PER split sub-group inside a
shared loop instead of once outside it, since each sub-group now has its
own lookup key). If you add a 6th lane or a new split sub-group, follow
this pattern, not the code-keyed one described in any stale comment you
might still find.

Critical detail preserved from the first round: the STRUCTURAL code (the
hardcoded default, e.g. `"G0001"`) must keep flowing into every internal
lookup that depends on it regardless of the override key-scheme change -
`PP_COMMODITY[code]`, `_charge_codes_for(code)` in lawc.py, `code ==
COMMODITY_ISC_MAIN[0]` in laec.py, `HNSLO_ROUTE_NOTE_BY_COMMODITY.get(
structural_code)` in lawc.py's `_build_rates_row` (which takes BOTH
`structural_code` and `output_code` as separate parameters for exactly
this reason - got this wrong once, conflating them, before splitting the
parameter). Only the code written into the final row's
`commodity_group_code` field goes through the override.

**`type` is now always `"C"`** on every `RatesRow`, every lane -
previously SAF hardcoded `"C"` but EAF/CSE/LAEC/LAWC all hardcoded `None`
to match their own ground truth (which genuinely does leave it blank in
all 4 of those samples - verified, not a guess). This is now overridden
as a deliberate business rule; each of those 4 lanes' test files excludes
`type` from the golden-diff (`TYPE_OVERRIDE_IGNORE` / `RATES_IGNORE_FIELDS`
per file) with a comment explaining why, rather than silently dropping
the check.

**UI**: Step 3's commodity editor (`step3_customize.py`) has an editable
`code` column (previously `disabled=["code", "description"]` made it
read-only) alongside description, both keyed by `default_description` per
the above. To keep multi-round editing correct - editing overrides,
applying, then coming BACK to Step 3 to edit again - the editor is seeded
from `WizardState.default_commodity_groups` (captured once at Step 2's
first parse, see §3.5 above), never re-derived from the already-
customized `row_sets`. Also fixed a latent bug while here:
`step1_upload.py`'s "new file" reset block had a comment claiming it
resets `profile` but the code never actually did - now it does
(`state.profile = MappingProfile()`), which matters because
`default_commodity_groups` is captured using whatever `state.profile`
holds at that first parse, and a stale profile from a previous file would
have corrupted that snapshot.

## 3.7 Multi-file upload + user-controlled commodity-group order

Two more features added in the same later session, both direct user
requests, both verified against a real 2-file split of the CSE sample
(byte-for-byte identical `to_opus_rows()` output whether parsed from the
one combined sample workbook or from 2 files split and re-merged).

**Multi-file upload** (`excel_io/merge.py`, new): the user pointed out
CSE genuinely ships as 2 real-world files - a main "Tier 1" file and a
separate "...for VELAG and VEPBL" file (VELAG/VEPBL are Venezuelan port
codes, La Guaira/Puerto Cabello - this is exactly the "CSE VE" sheet,
confirmed by reading its raw cells). `merge_workbooks(list[Workbook]) ->
Workbook` copies every sheet from every uploaded workbook into one target
workbook (cell values + font + fill, since `excel_io/style_utils.py::
is_excluded()` only checks those two - and merged-cell ranges, since
`header_grid.py::flatten_pod_header()` needs them), raising
`DuplicateSheetError` if two uploads both contain a sheet of the same
name rather than silently overwriting. Critical detail: copy EVERY cell
in the source's used range, including blank ones, not just non-None
ones - `dst_ws`'s own `max_row`/`max_column` are inferred from which
cells were ever touched, and `lawc.py`'s single-column sheet parser
(`_parse_single_col_sheet`) iterates up to `ws.max_row`/`max_column`
directly, so a shrunken destination range would have silently truncated
it. Wired into `step1_upload.py` (`st.file_uploader(...,
accept_multiple_files=True)`, `WizardState.upload_names: list[str]`
replacing the old singular `upload_name`/`upload_bytes`) and `cli.py`
(`input` argument is now `nargs="+"`) identically - both just load N
workbooks and call `merge_workbooks()` before anything else happens, so
every downstream step (classification, all 5 parsers, the writer) is
completely unaware more than one file was ever involved.

**User-controlled commodity-group order**: Step 3's commodity table has
an **Order** column - Streamlit 1.62's `st.data_editor` has no native
row-drag-reorder capability (checked directly: no `reorderable`-style
parameter, no `st.*` widget name containing "order"/"sort"/"drag"), so a
plain editable integer column doing the same job (sort by it on Apply) is
the practical equivalent, not a literal drag interaction. New
`MappingProfile.commodity_group_order: list[str]` holds the FINAL
(post-override) commodity_group_description values in the user's chosen
sequence. Implementation:
- `parsers/common/ordering.py::reorder_by_group()` (generic: stable-
  reorders any list into blocks matching a group-order list, keyed by an
  arbitrary `key_fn`; unlisted groups keep their first-seen relative
  order, appended after every explicitly-ordered one) and
  `reorder_row_set()` (applies it to `rates`/`rates_port_port`, keyed by
  `commodity_group_description`, and to `cmdt_notes`, keyed by a NEW
  `CmdtNoteRow.group_description` bookkeeping field - added specifically
  because a CMDT NOTE block's rows don't otherwise carry any tag
  identifying which commodity group they belong to, and reordering by
  parent-row `contents` text would be fragile. Confirmed safe to add:
  `schema/opus_columns.py::CMDT_NOTE_ROW_FIELDS` is an explicit hardcoded
  list the writer zips against, doesn't include this field, so it never
  reaches the written file).
- `parsers/common/commodity.py::build_notes_by_description()` now tags
  every row it returns (parent AND children, not just the parent) with
  `group_description=spec.description`.
- `ui/parsing.py` (new): `run_parser(parser, workbook, profile)` wraps
  `parser.run_multi()` + applies `reorder_row_set()` if
  `profile.commodity_group_order` is set. Both `step2_preview.py` and
  `step3_customize.py` call this instead of `parser.run_multi()`
  directly, so the ordering step can never be accidentally applied in one
  place and forgotten in the other.
- **Known scope limit, deliberately not solved**: on LAWC specifically,
  `OPUS RATES PORT-PORT` uses a DIFFERENT `commodity_group_description`
  than `OPUS RATES` for the same group (the `PP_COMMODITY` remap from
  §3.6's LAWC work, e.g. RATES says "China_TWN_SIN_HKG_KR Dry", PORT-PORT
  says "FAK - China_TWN_SIN_HKG Dry" for the identical rows) - since
  `commodity_group_order` is built from RATES-level descriptions, it
  simply doesn't match anything on LAWC's PORT-PORT rows, so that one
  sheet on that one lane keeps its default order regardless of what the
  user sets. RATES and CMDT NOTE - the two sheets a user actually forms a
  mental model of "order" from - are unaffected everywhere, including
  LAWC. Fixing this fully would mean threading the RATES-level
  description through `PP_COMMODITY`'s remap as a second tag, which
  seemed like a lot of added complexity for a sheet nobody visually scans
  block-by-block the way they do RATES/CMDT NOTE - flagged here rather
  than silently left unmentioned.

## 3.8 CSE "CSE" grid sheet: dynamic row offset + real destination/column extent

Bug report (real user file, not the bundled sample): after §3.7's multi-file
upload, uploading the real 2-file CSE filing ("...Tier 1.xlsx" +
"...Tier 1 for VELAG and VEPBL.xlsx") produced an output with `OPUS RATES`
rows for every commodity group EXCEPT the main "CSE" one - "CSE VE" and
"CSE (MAOVLD)" were both present and correct, "CSE" was silently zero rows.

Root cause, found by diffing the real file's "CSE" sheet against the
bundled sample cell-by-cell: the real file has an extra note row
("For PAMIT,PACFZ,DOCAU : INCL THD...") inserted above the
`SERVICE SCOPE = CSE` anchor line that the sample doesn't have, shifting
every header/data row below it down by exactly 1. `GridSheetConfig`'s
row numbers were hardcoded against the sample's (unshifted) layout, so the
parser read the real file's POD-NAME row as the POD-CODE row, tried to
match `container_map.suffix_for("Manzanillo, PA")` against D2/D4/D5 labels,
got `None` for every column, and produced zero `GridRow`s for the whole
sheet - the "silently skip a mismatched destination" `continue` (correct
behavior for *some* blank/malformed columns) was masking a *systemic*
row-offset failure, not a handful of bad cells. Verified this shift is
sheet-specific: "CSE (MAOVLD)" in the same real file matches its hardcoded
rows exactly (no anchor line precedes it in either the sample or this real
file) — only "CSE"/"CSE VE" carry the `SERVICE SCOPE = CSE` marker.

The fix (`parsers/cse.py`), scoped to the 3 `GRID_SHEETS` entries only:
- `_find_anchor_row()`/`_resolve_grid_config()`: search rows 1-20 for the
  literal text `"SERVICE SCOPE = CSE"`; if found at a different row than
  the config's built-in offset implies, shift `pod_code_row`/
  `container_label_row`/`data_min_row` by the same delta before parsing.
  Sheets without the marker (`CSE (MAOVLD)`) fall through unchanged. This
  fixes both the sample (anchor already at its expected row -> no-op) and
  the real file (anchor one row lower -> shifts everything by 1) with one
  code path, instead of a second hardcoded row set.
- `GridSheetConfig.data_max_row` removed; `_parse_grid_sheet()`'s data-row
  loop is now open-ended, stopping at the first row with a blank column-B
  origin cell instead of a fixed row ceiling. Verified this exactly
  reproduces the old hardcoded boundary on all 3 sample sheets (a blank row
  always separates real data from the trailing "Free line detention..."
  footnote table both there and in the real file) while correctly picking
  up the real file's full data block without truncating it.
- `max_col` is now `max(cfg.max_col, ws.max_column)` instead of the sample's
  fixed value - the real file has 5 more destination columns
  (BBBGI/AWORJ/CWWIL/TTPTS/DOHAI) past the sample's rightmost column.
  Safe because `header_grid.flatten_pod_header()` already skips any column
  whose container-label cell is blank, so widening the scan range can only
  pick up real header columns, never garbage.

Verified against the real 2-file upload directly (not just golden tests):
"CSE" now produces 2484 `OPUS RATES` rows alongside "CSE VE" (276) and
"CSE (MAOVLD)" (276); spot-checked that all 5 of the real file's extra
destination codes appear in `OPUS RATES PORT-PORT` (4 of 5 have actual
rate data - the 5th, TTPTS, has genuinely blank cells for every origin in
the raw sheet, correctly excluded, not a parsing gap). All 6
`test_parsers_cse.py` golden tests against the bundled sample still pass
unchanged.

## 3.9 Withdrawn-location exclusion (struck/blacked-out origin+destination cells)

User-requested feature, general across every lane: "if a location is
supposedly not included, e.g the cell is blacked, a word has strikethrough
(the word/phrase should be deleted)." Exclusion styling (strikethrough
font, or a fill color in `ExclusionConfig.blackout_rgbs`) was already
applied per-rate-cell everywhere (`is_excluded()`, one call per lane's data
loop) - a struck rate for one origin/destination pair was correctly
dropped. What was missing: the SAME styling convention applied to a
location's own name/code cell (marking the whole origin or destination as
withdrawn, not just one of its rates) wasn't checked anywhere - a withdrawn
POR/POD would still be parsed and filed as if it were live.

Two additions, both in `parsers/common/`, wired into every lane:

- `exclusion.py::location_is_excluded(cells)` (new): true if ANY of the
  passed cells is struck/blacked. Deliberately `any()`, not `all()` - a
  location is usually only 1-2 identifying cells (name + code), and a
  trader marking just one of them (typically the code) is the realistic
  case; requiring both marked would silently keep a withdrawn location
  whenever the formatting wasn't applied to every cell. (The pre-existing
  `row_is_excluded()`, which does require every cell excluded, stays as-is
  for its original purpose - a row of many independent rate cells, where
  one stray strike shouldn't drop the whole row. Both are exported; each
  lane's origin-row check uses `location_is_excluded()`, never
  `row_is_excluded()`.)
- `header_grid.py::flatten_pod_header()`: gained an `exclusion_config`
  param and now checks the POD label cell itself (the merged range's
  top-left cell, same cell already read for the label text). A hit resets
  `current_pod` to `None` via a private `_EXCLUDED` sentinel distinct from
  "no label at this column" (plain `None`) - so every column under a
  struck/blacked POD gets dropped exactly like the existing "blank
  container label, no fallback" skip path, without touching the merged-
  range walk itself. One shared fix here covers every lane that calls this
  function (SAF, CSE's 3 grid sheets, LAEC's DRY+ingauge sections, LAWC's
  3 dry sections + OOG).

Per-lane origin-row loops all gained the same one-line guard right after
their existing blank-cell check, checking the location's name cell
(column 1, verified consistent across every lane's raw layout) together
with whichever cell actually carries the code used as `origin_code_raw`:

- `cse.py`: `_parse_grid_sheet()`, `_parse_nor_sheet()` (also filters its
  `dest_codes` dict comprehension, which reads destination codes directly
  rather than through `flatten_pod_header`), `_parse_ingauge_sheet()`.
- `laec.py`: `_parse_grid_section()` (covers both DRY and, via
  `INGAUGE_SECTIONS`, the in-gauge sheet - same method), `_parse_nor_section()`
  (same direct-dict-comprehension `dest_codes` fix as CSE's NOR sheet).
- `lawc.py`: `_parse_dry_section()`, `_parse_single_col_sheet()` (covers
  both Reefer and NOR - same method, plus its own direct `dest_codes` dict
  built by scanning columns rather than `flatten_pod_header`), `_parse_oog_sheet()`.
- `saf.py`: single free-text origin column (no separate code cell, no
  meaningful "name" column to pair it with - column A there is a sparse
  region-grouping label, not a per-row identity) - checks `is_excluded()`
  on the one origin cell directly rather than `location_is_excluded()`.
- `eaf.py`: same single-column origin check as SAF, PLUS a sheet-level
  check - EAF's raw layout is one destination per whole sheet (not a
  multi-POD header), so a struck/blacked destination cell now empties
  `destination_text` and skips parsing any rate cells for that entire
  sub-lane sheet.

Deliberately NOT touched (out of scope for this pass - auxiliary surcharge
tables with a different shape than a POR×POD rate grid, not verified
against any real-world "withdrawn location" example): LAEC's ECSA Add-On
sheet, and the CSE/LAEC-shared Yangtze ARB Add-on sheet
(`common/yangtze_arbs.py`).

Tested with synthetic in-memory workbooks (no bundled sample exercises
this styling on a location-identity cell, only on rate cells) -
`tests/test_header_grid.py` (POD-level strike + blackout) and
`tests/test_parsers_cse.py` (`test_cse_grid_sheet_skips_withdrawn_origin`,
`test_cse_grid_sheet_skips_withdrawn_destination`). All existing golden
tests against the 5 bundled samples still pass unchanged - none of them
have a struck/blacked location-identity cell, so this is purely additive.

## 3.10 Wizard session cache could serve stale parse results

Reported as "§3.8's CSE bug is still there" after §3.8 had already fixed
it - the parser was correct (verified end-to-end: merge → classify →
parse → Step 3 apply → writer, in the user's own upload order, produced
CSE=2484 rows in the written workbook), but the running Streamlit app was
still showing the pre-fix result. Not a parsing bug at all; a caching one.

Two independent caches in the wizard, both keyed too weakly:

- `step2_preview.py` parses only `if state.row_sets is None`, and nothing
  else ever sets it back to `None` except a change of upload. So once a
  session has parsed, ANY later change that should alter the parse but
  doesn't change the upload (a parser code change, most obviously) is
  invisible - the wizard keeps replaying the cached rows through Steps
  3/4 and into the downloaded file.
- `step1_upload.py` compared `[f.name for f in uploaded]` against
  `state.upload_names` to decide whether to invalidate. Re-uploading the
  same filenames - which is exactly what a user does to "try again" -
  compares equal, so the cache (including `state.workbook` itself) was
  kept and nothing re-ran.

Fixes:
- `WizardState.upload_key` (new): a sha256 over each file's name AND its
  bytes. `step1_upload.py` now invalidates on this fingerprint instead of
  on names alone, so re-uploading an edited file of the same name
  correctly re-parses. (Also removes a latent bug of its own: an edited
  same-name file previously produced output from the PREVIOUS file's
  contents, silently.)
- `step2_preview.py` gained a "↻ Re-parse from source" button that clears
  `row_sets`/`default_commodity_groups`/`output_bytes` and the override
  profile, then reruns. The profile reset is required, not incidental:
  `default_commodity_groups` is re-snapshotted from the next parse and
  must be taken override-free (same reason Step 1's reset clears it).

Note for anyone debugging a "the fix didn't work" report against this UI:
check whether the app was restarted AND the parse recomputed before
concluding the parser is wrong. Streamlit reloads changed local modules on
rerun, but session_state (and therefore `row_sets`) survives that reload.

These were extracted as real reuse emerged — don't rebuild lane-specific
versions of these for LAWC or future lanes without checking here first:

- `commodity.py` — `resolve_commodity_code()`/`resolve_commodity_description()`
  (user-override lookups) and `CommodityNoteSpec`/`merge_note_specs()`/
  `build_notes_by_description()` (per-sheet default descriptions +
  merge-on-match CMDT NOTE building) - all described in §3.6.
- `header_grid.py` — flattens a merged/fill-down POD (or origin-group)
  header row into `{label: {container_label: col_idx}}`. Works for both
  "name row + code row" (SAF) and "code-only row" (CSE/LAEC) layouts.
- `exclusion.py` — wraps `excel_io.style_utils` for strikethrough/fill
  detection (feature requirement: blacked-out/struck cells excluded).
- `container_map.py` — loads `config/container_maps/{lane}.yaml`
  (raw container label → OPUS 20/40/40hc slot, plus lane-constant
  Prefix/CGO TYPE). Supports an optional reefer container label (SAF's
  `RD5` → a separate Prefix "R" row, not a 4th slot).
- `group_codes.py` — loads `config/group_codes/{lane}.yaml` (regional
  origin-group shorthand like `FEBP`/`WPRD` → member-code list). **Curated
  YAML, not parsed from raw footnote text** — the footnote text is
  unreliable (stale membership, typos; verified case-by-case against each
  lane's own ground truth).
- `cmdt_notes.py` — builds the `OPUS CMDT NOTE` parent+children rows from
  a raw "Rate structure: incl. X/Y/Z" line. Has grown several per-lane
  knobs as contradictions surfaced between lanes (see §6) —
  `sequential_charge_seq`, `sort_text_names`, `charge_code_names_override`,
  `trailing_oft_row`. **Read this file's docstrings before assuming any
  one lane's convention generalizes.**
- `ordering.py` — `group_by_destination()`, stable-groups output rows so
  the generated sheet reads as organized blocks (user-requested: all
  same-prefix rows together, destinations grouped within each block).
  `reorder_by_group()`/`reorder_row_set()` (§3.7) - the user-controlled
  commodity-group ordering feature.
- `yangtze_arbs.py` — parses the "Yangtze ARB Add-on" sheet (inland China
  origins via barge/rail) into `OPUS ARBS` rows. Shared by CSE (verified
  against ground truth) and LAEC (same logic, no ground truth to check —
  see §6).

`schema/opus_rows.py::explode_rates_row()` (grouped `RatesRow` →
`RatesPortPortRow` list) blanks `cmdt_seq` and `commodity_note` on the
exploded copies — confirmed these are grouped-view-only fields across
every lane examined so far.

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
block structure independently (a non-blank `contents` starts a block;
the following blank-`contents` rows belong to it - mirrors the writer's
own fill-down convention), keys blocks by the parent's `contents` text, and only
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

**Real bug found and fixed during implementation:** `reconstruct_blocks()`
originally detected a block's parent row via non-None `header_seq`/
`note_seq`. Those fields are writer-assigned at Excel-export time and are
always `None` on a freshly-parsed `OpusRowSet`, including on parent rows -
so comparing a fresh parse against a reference file (the feature's actual
primary use case) always reported every real CMDT NOTE/SPECIAL NOTE block
as "missing," even when content was identical. Neither this task's own
review nor Task 2's original review caught it, because every test
fixture in both tasks set `header_seq`/`note_seq` and `contents` together
- detection never actually depended on which field was checked. Fixed by
detecting via non-blank `contents` instead (already used for block
identity, verified non-blank on parents and blank on children in both
the fresh-parse and written-then-reread representations) - see commit
`4661ed6` and the regression test
`test_reconstruct_blocks_detects_parent_via_contents_when_seq_fields_are_none`
in `tests/test_audit_compare.py`. The Compare UI's `ignore_fields`
handling (added in the final-review fix wave) reuses each lane's
already-documented golden-test ignore sets - see
`RATES_IGNORE_FIELDS_BY_LANE` etc. in `mrg2opus/audit/compare.py`.

## 3.12 Per-group opt-out of DG (Dangerous Goods) duplicate rows

User-requested. Every lane (SAF, EAF, LAWC, LAEC, CSE) has a standing
filing convention, verified against each lane's own ground truth and
never derivable from the raw sheet itself: every base Dry (D/DR) row also
files an identical D/DG variant at the same rate. This was previously
unconditional - `MappingProfile` gains `skip_dg_generation: dict[str,
bool]`, keyed the same way as every other per-group override here (a
group's default, override-free description - see §3.6's key-convention
note). `True` suppresses that group's DG duplicate rows entirely (the
base D/DR rows are untouched); absent/`False` keeps the existing default
behavior, so every pre-existing golden test needed zero changes.

Each of the 5 lanes' DG-duplication code already lived inside a
`self.container_map.cgo_type == "DR":` guard, itself already inside a
per-group loop with the right description variable in scope
(`DEFAULT_COMMODITY_DESCRIPTION` for SAF/EAF's single/shared group,
`default_description` for LAWC/LAEC/CSE's per-sheet groups) - the fix is
one added clause per lane: `and not
config.skip_dg_generation.get(<description>, False)`.

**Known, accepted scope limit, same category as §3.6's LAWC PP_COMMODITY
note:** EAF's two sub-lanes (TZDAR, KEMBA) share one default description
("FAK") - already true for every other override in this profile, so this
toggle can't target just one sub-lane either. Confirmed with the user
before implementing; not a bug, matches existing precedent exactly.

UI: Step 3's commodity-groups table gains a "Skip DG" checkbox column,
collected into `skip_dg_generation` on Apply the same way every other
override column is (`ui/steps/step3_customize.py`).

Tests: one regression test per lane (`test_<lane>_skip_dg_generation_...`
in each `tests/test_parsers_*.py`), each confirming (a) the *default*
(unset) run still produces DG rows exactly as before, (b) setting the
toggle for one group removes only that group's DG rows and leaves its
DR rows untouched, and (c) for the 3 multi-group lanes (LAWC, LAEC, CSE),
a *sibling* group's DG rows are unaffected - proving the toggle is
correctly scoped per group, not a lane-wide switch.

## 3.13 UI quick wins: clickable step nav, friendlier errors, un-stale Compare results, full-list CSV downloads

User-requested, from a general "what could improve the UI" pass:

- **Clickable step navigator** (`ui/app.py`): the wizard's top step row
  (previously plain `st.markdown` text) is now a row of buttons, so
  jumping to any step - not just backward - works directly instead of
  requiring the page's own "Continue"/"Back" buttons. Safe because every
  step's `render()` already checks its own prerequisites before doing
  real work.
- **Friendlier error messages** (new `ui/errors.py::show_error()`): every
  `except Exception` site that used to dump raw exception text straight
  into `st.error(f"...: {exc}")` now shows a plain-language message with
  the raw exception tucked behind a collapsed "Technical details"
  expander - someone filing freight rates shouldn't have to parse a
  Python/openpyxl traceback, but the detail is still one click away.
  Applied in `step1_upload.py`, `step2_preview.py`, `step4_export.py`,
  and `compare_page.py` (2 sites).
- **Compare mode's results can no longer go silently stale**
  (`compare_page.py`): `CompareState.results_computed_for` snapshots
  `(lane_id, rates_mode, apply_known_gaps)` at the moment a comparison
  runs; if the visible controls no longer match that snapshot, a warning
  banner appears above the (still-shown) old results instead of letting
  them sit there looking current.
- **Full-list CSV download for truncated detail tables**
  (`compare_page.py::_render_detail_table()`): the on-screen Missing/
  Extra/Field-mismatches tables still cap at 50 rows (a large mismatch
  count would make the grid unwieldy), but anything beyond that is now
  always available as a full CSV download - previously the 51st row was
  simply invisible with no indication more existed.

**Bonus fix found while testing the clickable nav**: jumping straight to
Step 2 with nothing uploaded yet used to hit `_run_parser()` with a
`None` lane_id and show a generic "Parsing failed" - misleading, since
the real issue was "you haven't uploaded a file yet". This state was
unreachable before the nav became clickable (Step 2 was only ever
reached via Step 1's own button, which doesn't appear until a file is
loaded and classified) but is now a real path a user can hit. Step 2
gained an upfront `state.workbook is None or state.selected_lane_id is
None` guard with a clear "go back and upload" message, matching the
pattern Steps 3/4 already had for their own prerequisites.

## 3.14 Real filing sheet names + ROUTE NOTE (RN) sheet support

User-provided real MRG→OPUS ground-truth file pairs (`reference/1_MRGs` +
`reference/2_OPUS`, delivered 2026-08-26, tracked in `reference/for
hackathon TRACKER.xlsx`) revealed the app's own writer output sheet names
(`OPUS RATES`, `OPUS ARBS`, `OPUS CMDT NOTE`, `OPUS SPECIAL NOTE`) never
matched any real OPUS filing convention (`RATES`, `ORIGIN ARBS`, `CMDT
NOTE`, `SPECIAL NOTE`) - an internal naming convention this project had been
using since the start, confirmed via user clarification to need
correcting before final export. The same investigation surfaced a
previously-unmodeled third note-sheet type, ROUTE NOTE (`RN`), scoped to
a specific route pair rather than a whole commodity sequence or the whole
service - distinct from CMDT NOTE and SPECIAL NOTE, not merely the
existing `RatesRow.route_note` column. See the `project-opus-note-sheet-
taxonomy` memory for the full naming-drift story (SRCHG/SUR are CMDT NOTE
mislabeled; FREETIME/DEMDET are the same schema under a lane-specific
name).

- **Real filing sheet names** (`excel_io/writer.py::_sheet_names_for_
  suffix`): `rates`→`RATES`, `rates_port_port`→`RATES PORT-PORT`,
  `arbs`→`ORIGIN ARBS`, `cmdt_notes`→`CMDT NOTE`, `special_notes`→`SPECIAL
  NOTE`. Deliberately does NOT touch `schema/opus_columns.py`'s
  `SHEET_NAME_*` constants - those match the older, hand-prepared bundled
  `Sample MRGs with OPUS FORMATS/*.xlsx` fixtures every `test_parsers_
  *.py` golden test reads from directly by file, and changing them would
  break every lane's regression suite for no reason (those static files'
  own sheet names don't change just because a Python constant does). The
  app's Step 2/4 UI labels (`SHEET_LABELS` in `step2_preview.py`) are
  unaffected too - purely in-app display text, not the exported file.
  `compare_page.py`'s Compare mode was updated to look for these same
  real names in an uploaded reference file.
  **Correction (same day):** initially built as `O.ARBS` (verified
  against LAEC's real output) before CSE's own real reference pair
  arrived and showed `ORIGIN ARBS` (spelled out) instead - the same
  per-preparer naming drift as SRCHG/SUR, this time on ARBS. User's call:
  use the spelled-out `ORIGIN ARBS` everywhere, not lane-specific.
- **New ROUTE NOTE (`RouteNoteRow`) schema + writer support**
  (`schema/opus_rows.py`, `schema/opus_columns.py::RN_HEADER`/
  `RN_ROW_FIELDS`, `excel_io/writer.py::_write_route_note_sheet`): same
  shape as `CmdtNoteRow` plus `route_seq` and 10 trailing columns
  (Receiving Term through Premium) confirmed on the real RN sheet.
  `header_seq`/`route_seq`/`note_seq` are placeholder running numbers
  assigned at parse time (confirmed acceptable - OPUS renumbers these on
  import - same treatment CMDT NOTE's `header_seq`/`note_seq` already
  got elsewhere in this codebase), not reproductions of any real
  OPUS-assigned number. Unlike CMDT NOTE, real RN rows are header-only
  (no child charge-code rows).
- **LAWC wired up fully** (`parsers/lawc.py::_derive_route_notes`): every
  `RatesRow` with a non-null `route_note` (HNSLO's MAR/MX2, OOG's
  KCI/OH/OWOH/OW cases) gets a matching RN entry.
  - **Real bug fixed along the way**: `_oog_route_note` was returning the
    bare text `"IG"` for plain in-gauge equipment - ground truth (2
    independent newer reference files, FAK and Tier 1) shows this should
    be blank, and the KCI-flagged case should read bare `"...KCI"` with
    no `"(IG)"` suffix. The bundled `LAWC.xlsx` sample actually has the
    old `"IG"`/`"...KCI (IG)"` text baked into its own ground truth - a
    confirmed-stale artifact, not a second valid convention - so
    `route_note` is now excluded from that bundled sample's own
    RATES-matching tests (see their inline comments) and verified instead
    against the newer files directly (`test_lawc_route_notes_
    reference.py`).
  - **New real business rule, user-confirmed**: "REEFER DRY AS DANGEROUS"
    - Non-Operating Reefer (`LAWC NOR` raw sheet) cargo that's dangerous
      gets filed as `D`/`DG` (folded into G0004/S.E.A_JPN_SA_AU_NZ's
      regular dry-and-dangerous bucket) instead of `R`/`DG`, with this
      route note explaining why. This is why NOR's own commodity code
      shows as G0004 in the real reference files rather than the G0001
      default - not a parser bug, since commodity codes are user-
      customizable per lane already. `lawc.py`'s reefer/nor loop now
      builds this DG-duplicate for NOR specifically (never Reefer -
      `REEFER_CONFIG`'s cgo_type `"RF"` never matched the DG-duplication
      rule anyway), combining with any pre-existing route note (e.g. the
      `AX3` vessel-lane note) via `" | "`.
  - **Known, deliberately out-of-scope gap found along the way**: the new
    reference files' ISC/SEA raw sheets have real HNSLO destination rows
    this parser fails to detect in some cases (MAR route-note count short
    190 vs 196; MX2 entirely 0 vs 88 for the FAK reference file) -
    unrelated to route notes, confirmed pre-existing, and explicitly not
    chased down as part of this task (matches the already-deferred "LAWC
    Tier 1/FAK real-file fidelity" item). `test_lawc_route_notes_
    reference.py` checks route-note category *counts* rather than a full
    RN-sheet equality specifically to sidestep this gap without hiding
    regressions in what this task actually built.

New tests: `tests/test_writer.py` (sheet renaming + RN write/omit
behavior), `tests/test_lawc_route_notes_reference.py` (route_note/RN
correctness against the real `reference/2_OPUS/15_LAWC FAK` file pair,
skipped if that directory isn't present in the checkout).

## 3.15 Test suite rebuilt onto reference/ ground truth; old bundled samples deleted

2026-08-26 (later the same day as §3.14): the user deliberately deleted
`Sample MRGs with OPUS FORMATS/*.xlsx` and `MRGs RAW SAMPLES/*.xlsx` for
good (see feedback-reference-folder-convention memory) - every
`test_parsers_*.py` golden test read from these directly, so the whole
suite needed rebuilding onto `reference/1_MRGs`/`2_OPUS` pairs instead.

- **CSE, EAF (KEMBA only), LAEC, LAWC**: migrated to real `reference/`
  pairs. **SAF**: `test_parsers_saf.py` deleted outright - no `reference/`
  data exists for SAF at all (not even in the tracker).
- **`ARBS` naming corrected mid-migration**: CSE's real output uses
  `ORIGIN ARBS` (spelled out), not the `O.ARBS` §3.14 originally shipped
  (verified against LAEC only) - user's call: use the spelled-out name
  everywhere, not lane-specific. Fixed in `writer.py`/`compare_page.py`.
- **Real bugs found and fixed along the way** (each only surfaced because
  this was the first time these real files were compared field-by-field):
  - `eaf.py`'s sub-lane sheet detection (`SUBLANE_SHEET_RE`) required a
    literal `"EAF "` prefix (`"EAF TZDAR"`) - real EAF files use bare
    `"TZDAR"`/`"KEMBA"` sheet names and were completely unrecognized
    (zero rows parsed, not just a field mismatch). This would have hit
    real users uploading real EAF files through the UI today.
  - `schema/charge_codes.py::INDIVIDUAL_CHARGE_CODES` excluded `"BAF"`
    (Bunker Adjustment Factor) based on the old bundled EAF sample. User-
    clarified: their SOP tells human filing agents not to file BAF - a
    special case for people, not a filing-format rule - so this tool
    should reproduce the raw MRG's own "Includes" text as-is, BAF
    included. Confirmed against real KEMBA ground truth (2 independent
    weeks). Added, with `CHARGE_CODE_NAMES["BAF"]`.
  - `eaf.py`'s `PREPAID_LINE`/`trailing_oft_row` feature (added an "Ocean
    Freight to be Prepaid, payable at -1 by -1." line + an "OFT" child
    row) never matched either real KEMBA file - the `-1 by -1` text was a
    dead giveaway this was never fully implemented. Removed entirely,
    along with the now-dead `trailing_oft_row`/`extra_content_lines`
    parameters on the shared `parsers/common/cmdt_notes.py::
    build_cmdt_notes()`.
  - `cse.py`'s SPECIAL NOTE (PSA surcharge) validity window and its
    "Valid from X until Y" Contents text were hardcoded to one specific
    week's dates (`PSA_VALIDITY_START = date(2026, 5, 22)`, "March 22,
    2026 until August 31, 2026") as if a "standing policy date" - a real
    reference file (different week) shows it should just mirror the main
    filing's own `data.validity_start`/`data.validity_end` dynamically.
    Fixed via new `_psa_contents()` helper.
- **Follow-ups flagged, not fixed** (each would need its own investigation
  - do not assume any of these are simple):
  - ~~CSE's 2-file upload (`FAK.xlsx` + `...for VELAG and VEPBL.xlsx`)
    crashes with `DuplicateSheetError`~~ - **FIXED, see §3.18.** (Was
    spawned as a separate task, `task_243523c4`; fixed directly instead.)
  - LAEC: 210 real rows (all `DG`/`D`, all Argentina destinations e.g.
    `ARLPG`/`ARZAE`/`ARUSH`) are completely missing from what the parser
    generates - likely a DG-duplicate rule tied to the "ECSA Add-On" raw
    sheet, not implemented. Excluded from `test_parsers_laec.py` via
    `_is_known_missing_gap`.
  - LAWC: the SEA grid's Central/South America destinations (304 rows -
    152 DR + their DG-duplicates) are missing for this specific FAK
    week/file - broader than the HNSLO-only gap §3.14 first found via
    route-note counts. Excluded via `_KNOWN_MISSING_DESTINATIONS` in
    `test_parsers_lawc.py`. Also: NOR needs a THIRD DG variant for PEPAI
    specifically (plain `R`/`DG`, no route note) that isn't implemented
    either (12 rows, separately excluded).
  - **Fixed (same day, follow-up)**: LAWC's own hardcoded
    `MAIN_CHARGE_CODES`/`OOG_CHARGE_CODES`/`ISC_CHARGE_CODES`/
    `SEA_CHARGE_CODES` lists (unrelated to the shared `charge_codes.py`
    fixed above - LAWC doesn't go through `INDIVIDUAL_CHARGE_CODES` at
    all) were also missing `"BAF"`. Confirmed present in all 5 of
    reference/2_OPUS/15_LAWC FAK's real SRCHG blocks; added to all 4
    lists (placed first, matching the 2 groups directly observed with
    BAF as the first child row - see project-tool-mirrors-mrg-not-human-
    sop memory for why this isn't a stale-sample issue).
    `test_lawc_cmdt_note_merges_when_descriptions_match` stays removed
    regardless (its 25-vs-35 count mismatch has other, unresolved causes
    beyond just BAF).
  - CSE and LAEC's CMDT-NOTE-merge-by-description behavior (previously
    verified against the old bundled samples) doesn't reproduce either
    real reference file's actual block count/content - both merge tests
    were removed rather than reverse-engineered under time pressure.
- **Sheets that don't exist in real output**: neither CSE's, LAEC's, nor
  LAWC's real reference files have a `RATES PORT-PORT` sheet at all (the
  old bundled samples did) - those ground-truth tests were removed, not
  adapted, since there's nothing to compare against.
- **Cross-cutting fixture cleanup**: `test_mrg_upload.py`,
  `test_header_grid.py`, `test_registry.py` used `SAF.xlsx` purely as
  "some real xlsx" for upload/header-grid/classification mechanics
  unrelated to SAF specifically - repointed at LAEC/CSE `reference/`
  files instead (with matching assertion updates, e.g. lane_id/sheet
  names). `tests/golden.py` (dead once every lane stopped reading from
  it), `tests/conftest.py` (only fixture was the now-deleted sample
  directories), and `tests/test_compare_engine_regression.py` (redundant
  with the migrated per-lane tests, which now exercise the same
  production `audit.compare` functions directly) were deleted outright.

## 3.16 Filing-wide charge-code exclusion ("Special instructions")

Follow-up to §3.15's BAF fix: the user clarified BAF is excluded from
some real filings not because ground truth was stale, but because their
own SOP tells human filing agents not to file it for certain accounts
(BAF is an oil surcharge, functionally the same as OBS; their Hong Kong
account's RFAs don't apply it) - a human-only special case the TOOL
should still be able to reproduce the raw MRG's own text for by default,
with an opt-in override for exactly this kind of account-specific
exclusion. See project-tool-mirrors-mrg-not-human-sop memory.

- **`MappingProfile.excluded_charge_codes: list[str]`** (new field,
  `presets/models.py`) - filing-wide (not per-commodity-group, user's
  explicit choice), general-purpose (any charge code, not BAF-specific).
- **`parsers/common/cmdt_notes.py::build_cmdt_notes()`** gained an
  `excluded_codes: frozenset[str]` parameter, applied before anything
  else - an excluded code never appears in the "inclusive of" text or
  gets its own child row. `build_notes_by_description()` already forwarded
  arbitrary kwargs, so no change needed there.
- Wired into all 5 lanes' `to_opus_rows()` (`cse.py`, `eaf.py`, `laec.py`,
  `lawc.py`, `saf.py`) - every `build_cmdt_notes`/`build_notes_by_description`
  call site now passes `excluded_codes=frozenset(config.excluded_charge_codes)`.
- New "Special instructions" section in Step 3 (`ui/steps/step3_customize.py`):
  one comma-separated text input, filing-wide, feeding the new profile field.
- Tests: `tests/test_cmdt_notes.py` (unit-level: exclusion removes both
  the child row and the text mention; excluding every code yields no
  CMDT NOTE at all; default excludes nothing) and a new end-to-end test
  in `test_parsers_eaf.py` confirming the profile field reaches a real
  lane's output.
- Verified live in the browser (Step 3 → Customize, real LAEC file):
  the field renders, accepts input, and Apply & Continue flows through
  to Export with no errors.

## 3.17 Sequential default commodity group codes (G0001, G0002, ...)

User-directed (2026-08-27): every distinct commodity group now gets its
own unique output code by default, instead of silently sharing whatever
structural code a lane's parser happens to use internally for unrelated
joins (e.g. LAWC's main dry grid/Reefer/LAWC NOR all default to the SAME
internal `"G0001"` - see `PP_COMMODITY`/`HNSLO_ROUTE_NOTE_BY_COMMODITY`
in `lawc.py`, which still key off that shared internal code and were
NOT touched). Numbered G0001, G0002, G0003, ... in the order groups are
first encountered while parsing.

- **`ui/commodity_utils.py`**: `distinct_commodity_groups()` now returns
  groups in first-encounter order (was `sorted()` by code+description -
  a behavior change to Step 3's default row order too, before the user
  sets their own "Order" values). New `assign_sequential_default_codes()`
  builds the `{description: "G000N"}` mapping.
- **`ui/steps/step2_preview.py`**: right after the first (override-free)
  parse and snapshotting `default_commodity_groups` (unchanged, still
  override-free - the sequential mapping is computed FROM it, not
  before), auto-seeds `state.profile.commodity_code_overrides` with the
  sequential mapping - only when the profile has no code overrides yet
  (so this never clobbers a loaded preset), then **re-parses once more**
  so `row_sets` (this preview, and Export if the user changes nothing
  further) reflects the new codes immediately, not just Step 3's editor.
- Internal structural codes are completely untouched - every parser's
  own internal joins keep working exactly as before; only what
  `resolve_commodity_code()` resolves to BY DEFAULT changed, via the
  existing override mechanism (no parser file was modified).
- Verified live in the browser against LAWC's real reference file:
  Step 3's "Parsed default code" column (disabled) still shows the
  original structural codes (G0003 ISC, G0004 SEA, G0001×3 for Main/
  Reefer/NOR, G0002 OOG); "Code (yours)" now shows unique G0001-G0006.
- Tests: new `tests/test_commodity_utils.py` (order preservation +
  unique-code assignment).

## 3.18 Fixed: CSE's real 2-file upload crash

Follow-up from §3.15/§3.16: uploading CSE's two real files together (the
main FAK file plus a separate "...for VELAG and VEPBL" Venezuela file)
crashed with `DuplicateSheetError` - both files share 4 sheet names
("CSE", "DG surcharges", "Yangtze ARB Add-on", "Free Time") and
`excel_io/merge.py` had no rule to reconcile them, despite its own
module docstring describing exactly this real-world 2-file pattern.

- **`excel_io/merge.py::merge_workbooks()`** now accepts an optional
  `names` list (the original filenames, same order as workbooks). When a
  colliding file's own name mentions "VELAG"/"VEPBL" (case-insensitive -
  the confirmed real pattern, not a content-based guess): its `"CSE"`
  sheet is renamed to `"CSE VE"` (the name `cse.py::COMMODITY_VE`
  expects) instead of raising, and its duplicate `"DG surcharges"`/
  `"Yangtze ARB Add-on"`/`"Free Time"` sheets are silently dropped (the
  main file's copies win - confirmed either byte-identical or a
  formatting-only variant, not Venezuela-specific data). Any OTHER
  collision (no `names` given, or a name that doesn't match the pattern)
  still raises exactly as before - `tests/test_excel_io_merge.py`'s
  original `test_duplicate_sheet_name_raises` is unchanged.
- `ui/mrg_upload.py::load_and_classify()` and both its callers
  (`step1_upload.py`, `compare_page.py`) now thread the uploaded
  filenames through to `merge_workbooks()`.
- **A second real bug found while verifying the fix**: `cse.py`'s "In
  guage guideline" sheet parsing had a hardcoded `INGAUGE_MAX_COL = 28`
  that silently dropped 2 real destinations (BRVLD, GYGEO) sitting past
  column 28 in this same real file - the same "real file wider than the
  verified minimum" class of bug already fixed once for the main "CSE"
  grid (`_resolve_grid_config`), just not applied to this sheet. Now
  widens to `max(INGAUGE_MAX_COL, ws.max_column)`.
- `tests/test_parsers_cse.py` now merges both real files (previously
  documented as only testing the main file alone) and asserts full
  parity - **0 missing, 0 extra, 0 field mismatches** against all 4898
  real ground-truth RATES rows.
- Verified live in the browser: uploading both real CSE files together
  no longer crashes, parses to exactly 4898 rows matching ground truth.

## 3.19 New lane: West Africa WAF

First brand-new lane built from `reference/` alone (no old bundled
sample ever existed for it) - a re-scan of `reference/` turned up 9
lane families with real, tracker-`Completed` MRG+OPUS pairs and zero
parser code (see project-mrg-lane-scope memory); this is the first.

- Single raw sheet ("Asia WAF MRG FAK rate guideline"): a 9-POD (Apapa,
  Tincan, Onne, Lekki, Tema, Abidjan, Lome, Cotonou, Dakar) x 3-container
  (D2/D4/D5 -> 20/40/40HC) grid over ~41 fixed Asia/SEA origins. Output
  scope is RATES + CMDT NOTE only (ground truth's own sheet is literally
  named "SRCHG" - the already-known CMDT-NOTE naming drift) - no ARBS,
  SPECIAL NOTE, or RN, confirmed against both real weekly ground-truth
  files.
- **New parser: `mrg2opus/parsers/waf.py`.** Origin/destination code
  resolution needed ZERO hardcoded per-lane port table - the existing
  Location Bank (`location_bank/fuzzy_match.py`) already resolved all 41
  Asia/SEA origins correctly (shared with other lanes), confirmed by a
  dry-run diff against the exact ground-truth code/description for every
  one before writing a line of the real parser. Only the 9 West Africa
  *destinations* were genuinely new - added to the Location Bank
  (`data/location_bank.sqlite3`) as `source="manual_override"` records,
  each with an exact alias matching the raw sheet's own POD label
  (Apapa->NGAPP, Tincan->NGTIN, Onne->NGONN, Lekki->NGLKK, Tema->GHTEM,
  Abidjan->CIABJ, Lome->TGLFW, Cotonou->BJCOO, Dakar->SNDKR).
- New container map `mrg2opus/config/container_maps/waf.yaml` (D2/D4/D5 ->
  20/40/40hc, no reefer).
- Raw origins use a parenthetical via-clause format never seen before
  ("Ganzhou (via Shekou)") - existing `split_location_text`'s via-regex
  only handles the "via X" (no parens) form used elsewhere, so `waf.py`
  has its own `_split_via()` for this lane.
- Every base D/DR row also files an identical D/DG duplicate at the same
  rate under its own commodity group ("<desc> - DG") - the raw sheet's
  separate per-IMO-class HAZ/PSA add-on tables (its own "*HAZ/PSA tariff
  table" blocks) are NOT reflected anywhere in either ground-truth file
  (no extra charge codes, no ARBS), so they're out of scope and
  deliberately not parsed at all.
- Two new charge codes discovered and added to `schema/charge_codes.py`:
  `CGD` ("CONGESTION SURCHARGE (D)") and `EPH` ("ELSEWHERE PAYMENT
  HANDLING FEE" - verbatim 3 literal spaces between the two words, copied
  exactly from ground truth text). The raw sheet's "Incl. BAF, HEA, EPH,
  BRS, LSF, CGD, OBS, MBS, EFS;" text is comma-separated (not the
  "/"-separated convention `cmdt_notes.py::parse_included_charge_codes`
  handles), so `waf.py` has its own extraction regex; BRS is correctly
  dropped (not in `INDIVIDUAL_CHARGE_CODES`, same already-known pattern
  as EAF's BRS).
- **Two real formatting quirks found only by diffing against ground
  truth, not guessable from the raw sheet alone:**
  - `Route Seq.` is a single running counter across the WHOLE 369-row
    commodity-group block (all 9 PODs x 41 origins in encounter order),
    NOT resetting per destination as most other lanes' route-seq-like
    fields might suggest - confirmed the Tincan POD block continues
    42, 43, ... rather than restarting at 1.
  - This lane's own ground truth spells every multi-part origin/
    destination name "CITY  SUBDIVISION" (double space, zero exceptions
    across 37 distinct names in both weeks) instead of the Location
    Bank's "CITY, SUBDIVISION" (comma) convention mined from other
    lanes' ground truth - `waf.py::_clean_description()` converts.
- Registered in both entry points (`cli.py`, `ui/app.py`)'s side-effect
  import list; added to `tests/test_registry.py`'s `ALL_LANE_IDS` and
  `SAMPLES` (confirmed 100% classification confidence, correctly
  disambiguated from the similarly-named but distinct, still-unbuilt
  "West Asia to West Africa" lane, whose sheet is literally named "WAF"
  with title "West Asia WAF FAK rate guideline" - close enough in
  wording that `sheet_name_patterns`/`title_keywords` needed to target
  this lane's actual full sheet name, "Asia WAF MRG FAK rate guideline",
  rather than a generic "WAF" substring).
- New `tests/test_parsers_waf.py`: both real weekly pairs, **0 missing, 0
  extra, 0 field mismatches** against all 738 real ground-truth RATES
  rows each week, plus a CMDT NOTE content match (folder 9's own SRCHG
  sheet carries an extra, verified-duplicate 18 rows belonging to the
  NEXT week - a copy-paste leftover in that one ground-truth file itself,
  not a parsing gap - documented and sliced off in the test rather than
  chased). `excluded_charge_codes`/`skip_dg_generation` wiring also
  covered.
- Verified live in the browser: full 4-step wizard (upload -> 100%-
  confidence classify -> 738-row preview -> sequential G0001/G0002
  commodity codes auto-assigned -> export screen) end to end on a real
  reference file.

## 3.20 CMDT NOTE child rows' RFA effective/expiry override

User clarified the exact business reason behind §3.19's documented gap
(CMDT NOTE child rows' Application Effective/Expires not matching West
Africa WAF's ground truth): the child dates aren't derived from the
weekly rate validity at all - they're the charge code's own RFA (Rate
Filing Agreement) window, a separate, usually longer-lived date pair a
human filer enters per account. Confirmed filing-wide scope (same choice
as §3.16's `excluded_charge_codes`, not per commodity group).

- `MappingProfile` gains `rfa_effective_date`/`rfa_expiry_date: date |
  None = None`. Unset (the default) keeps the pre-existing fallback
  behavior (children mirror the weekly rate validity) exactly as before
  this feature existed.
- `cmdt_notes.py::build_cmdt_notes()` gains matching `rfa_effective`/
  `rfa_expiry` params - each bound falls back to `validity_start`/
  `validity_end` independently when its override is `None`. Only CHILD
  rows are affected; the parent (`APP`) row always keeps the weekly rate
  validity window regardless, matching every ground-truth example seen.
  `build_notes_by_description()` already forwards `**build_kwargs`
  unchanged, so no change needed there.
- Threaded through every lane's own `build_cmdt_notes()`/
  `build_notes_by_description()` call site (cse.py, laec.py, lawc.py,
  eaf.py, saf.py, waf.py) - filing-wide, not lane-specific, even though
  it was only directly confirmed against WAF's ground truth so far.
- UI: `step3_customize.py`'s "Special instructions" section gains two
  optional `st.date_input` widgets ("RFA effective date"/"RFA expiry
  date", `value=None` when unset, confirmed renders as an empty
  yyyy/mm/dd picker in Streamlit 1.62).
- New tests: `test_cmdt_notes.py` (4 new cases - default fallback
  unchanged, override applies to children only, bounds are independent)
  and `test_parsers_waf.py::test_waf_rfa_override_matches_ground_truth_child_dates`
  (end-to-end: reproduces WAF's exact real ground-truth child dates,
  20260520-20261231, when the override is supplied). Verified live in
  the browser: both date pickers render, accept input, and the filing
  completes through Export with no error.

## 3.21 New lane: AUS NEA to AUEC FAK

Second brand-new lane, chosen from the same re-scan that found WAF -
three of the completed-but-unbuilt families shared WAF's "grid over
Asia/SEA origins" shape (see project-mrg-lane-scope memory); this one
turned out structurally richer than WAF in several ways.

- Three raw sheets, one shared 48-origin grid layout (same origins, same
  row range, same footnote city-group definitions) but different
  destinations: `ex NEA to AUBNE_AUMEL` (Brisbane/Melbourne, one combined
  POD) and `ex NEA to AUSYD` (Sydney) both feed ONE commodity group
  ("EX NEA TO AUEC") - confirmed sharing a single CMDT NOTE block in
  ground truth; `ex NEA to AUBNE on NZJ` (Brisbane via a different vessel
  operator) is a wholly separate SECOND commodity group ("EX NEA TO
  AUBNE ON NZJ"). Output scope is RATES + CMDT NOTE only (ground truth's
  own sheet is named "SUR" - the known naming drift).
- **New parser: `mrg2opus/parsers/auec.py`.** Origin resolution again
  needed almost no new Location Bank entries - of 48 origins, only 3
  new manual_override records (Guiyang->CNKWE, Shidao->CNSHD,
  Mawei->CNMAW) plus 2 defensive aliases for "Taizhou,  Jiangsu" (double
  space) pointing at the already-known CNTZO, needed because a bare
  comma-split would wrongly treat "Taizhou" and "Jiangsu" as two separate
  locations - `auec.py` now tries the whole (unsplit) origin text as one
  token FIRST, falling back to the normal comma-split multi-match only
  if that fails, specifically to handle this "City, Province" format.
- **New per-lane concepts not seen in any prior lane:**
  - A container type genuinely new to this project: "RAD" ("Reefer As
    Dry") - a physically reefer container filed as Prefix R, CGO TYPE
    DR (not RF), using the same 20'/40'HC-only rate shape as the RF
    (real reefer) columns. Confirmed via the raw sheet's own remark,
    "*not accepting DG in Reefer" - RAD rows never get a DG duplicate.
  - A rail-routed origin distinction: "Chengdu (via CNYTN by rail)" sets
    `origin_transmode="Rail"` (NOT `origin_term`, which stays "CY"
    either way) - only the literal "by rail" text triggers it; a plain
    "(via CNSHA)" without "by rail" just sets O.Via and leaves
    transmode blank. Easy to get backwards (did, initially - caught by
    the ground-truth diff).
  - Two named regional groupings, "PRDA*"/"PRDB*", that expand to 19
    Pearl-River-Delta cities each via a footnote definition elsewhere in
    the same raw sheet (rows 56-57) rather than being written out
    inline. Several of those 38 individual city names (e.g. "Longhua",
    "Sihui (Mafang)") are new enough, and ambiguous enough, that a fuzzy
    match mis-resolved some of them to unrelated existing codes with
    a real risk of silently wrong freight rates - `auec.py` hardcodes
    both groups' full code/description strings instead, copied verbatim
    from ground truth (confirmed byte-identical between both reference
    weeks, a stable standing grouping) rather than resolved per-city.
  - A per-charge-code POL (origin country) scope on CMDT NOTE child
    rows: ISL is only included for Taiwan-origin shipments; EFS is filed
    as TWO separate child rows scoped to Hong Kong and Korea
    respectively (both genuinely present, not a duplicate-row bug - same
    category as CSE's real duplicate THL row). None of this is
    recoverable from the raw sheet's own "Rate incl OBS, EFS (ex KR &
    HK), PSS, subject to ISL/..." remark text (which actually lists ISL
    under "subject to", not "incl") - the whole charge/POL list is
    hardcoded (`auec.py::INCLUDED_CHARGES`), verified against both real
    weeks. Needed a lane-specific CMDT NOTE builder
    (`auec.py::_build_group_cmdt_notes`) rather than the shared
    `cmdt_notes.py::build_cmdt_notes` helper, which has no notion of a
    per-child-row POL.
  - Route Seq. is a single running counter across the WHOLE commodity
    group - not just per destination (as WAF's own equivalent finding
    was), but across every destination AND container-type
    (DR/RF/RAD-DR/DG) block within that group, restarting at 1 only for
    the next commodity group.
  - Unlike WAF, the D/DG duplicate row stays under the SAME commodity
    group code/description as its D/DR parent (no "- DG" suffix split)
    - confirmed directly from ground truth; the two lanes genuinely
    differ here, not a WAF pattern to blindly reapply.
- One new charge code discovered and added to `schema/charge_codes.py`:
  `ISL` ("INTERNATIONAL SECURITY FEE AT LOCAL").
- Sheet-name-based detection needed to be specific (`ex NEA to
  AUBNE_AUMEL` / `ex NEA to AUSYD` / `ex NEA to AUBNE on NZJ` exactly)
  since the sibling, still-unbuilt "AUS NEA to AUWC" lane shares the
  identical raw template with different destination sheet names (`ex NEA
  to AUFRE` / `ex NEA to AUADL`) - confirmed 100% classification
  confidence with zero ambiguity in `test_registry.py`.
- New `tests/test_parsers_auec.py` (7 tests): both real weekly pairs, **0
  missing, 0 extra, 0 field mismatches** against all 348 real
  ground-truth RATES rows each week; CMDT NOTE content match including
  the POL scoping; a dedicated route_seq-continuity test; and
  `excluded_charge_codes`/`skip_dg_generation` end-to-end coverage
  (confirming skip_dg_generation is correctly scoped per commodity group
  - suppressing the main group's DG rows leaves the separate NZJ group's
  DG rows untouched).
- Verified live in the browser: full 4-step wizard (upload -> 100%-
  confidence classify -> 348-row preview -> customize -> export screen)
  end to end on a real reference file.
- Not yet verified: the sibling "AUS NEA to AUEC TIER 1" reference pair
  (folders 39/40) - same 3 sheets plus an extra "Tier 1 list" sheet this
  parser already ignores harmlessly, so it likely just works, but this
  wasn't explicitly checked.

## 3.22 New lane: AUS NEA to AUWC FAK

Third brand-new lane this session, deliberately NOT assumed to be
AUEC-identical despite sharing the "AUS NEA to Aus-coast" family name -
per [[project_mrg2opus_auec_lane]]'s own warning, it wasn't. Verified
against both real reference weeks: 0 missing/0 extra/0 mismatched rows
across all 248 RATES rows each week, plus exact CMDT NOTE parity.
Verified live in the browser end to end.

- Two raw sheets, "ex NEA to AUFRE" (Fremantle) and "ex NEA to AUADL "
  (Adelaide), sharing the same 49-origin grid and PRDA/PRDB footnote
  definitions AUEC uses - byte-identical, so `auwc.py` imports
  `PRDA_CODE`/`PRDA_DESCRIPTION`/`PRDB_CODE`/`PRDB_DESCRIPTION` straight
  from `auec.py` rather than re-transcribing them. Both destinations
  share ONE main commodity group (confirmed one CMDT NOTE block),
  matching AUEC's AUBNE_AUMEL+AUSYD pattern.
- **New parser: `mrg2opus/parsers/auwc.py`.** All 49 origins resolved
  via the existing Location Bank with ZERO new entries needed - the 4
  manual_override records added during the AUEC build (Guiyang, Shidao,
  Mawei, the Taizhou-Jiangsu alias) already covered this lane's overlap.
- **Two ways this lane genuinely differs from AUEC - confirmed via
  ground truth, not assumed:**
  - Reefer (RF) and "NOR" (Non-Operating Reefer - AUEC's equivalent
    container type is named "RAD", same concept: physically a reefer
    container filed as Prefix R / CGO TYPE DR) each get their OWN
    SEPARATE commodity group here (`FAK - NEA to AUWC (RF)` / `(NOR)`),
    not folded into the main group the way AUEC's RF/RAD are. 3 distinct
    CMDT Seq blocks in ground truth, not 2. NOR only ever populates the
    20' rate slot in this lane's raw sheet (its own "40'HC NOR" column
    is always blank) - not a bug, just this container type's data shape
    here, unlike AUEC's RAD which populates both 20'/40'HC.
  - CMDT NOTE child rows DO carry a per-code origin scope (ISL only for
    Taiwan; EFS filed as two separate rows, here scoped to Korea and
    Hong Kong respectively) - same hardcoded charge list and same
    business meaning as AUEC's POL scoping, but this lane stamps it on
    the **POR** column instead of POL (confirmed: every POL cell here is
    blank). Needed the same kind of lane-specific CMDT NOTE builder
    AUEC's did (`auwc.py::_build_group_cmdt_notes`), just targeting a
    different OPUS field - this was caught by the ground-truth diff on
    the FIRST attempt (reused a plain, non-scoped charge list expecting
    AUEC's shared-helper path to just work), a reminder that "no per-code
    scoping" and "different scoping field" both present as the SAME
    "POL always blank" signal until you check the neighboring POR/PO*
    columns too.
  - Route Seq. resets to 1 for each of the 3 commodity groups
    independently (DG continues the main group's own counter, same
    "one continuous counter per commodity group" rule as AUEC) - this
    one DID match the established pattern.
- No new charge codes needed (OBS/ISL/EFS/PSS all already known from
  AUEC).
- Sheet-name detection (`ex NEA to AUFRE` / `ex NEA to AUADL`) is
  already fully distinct from every sibling AU/NZ lane's own sheet
  names - confirmed 100% classification confidence with zero ambiguity
  in `test_registry.py`.
- New `tests/test_parsers_auwc.py` (8 tests): both real weekly pairs, **0
  missing, 0 extra, 0 field mismatches** against all 248 real
  ground-truth RATES rows each week; CMDT NOTE content match including
  the POR scoping; a route_seq-per-group test; a 3-distinct-groups test;
  and `excluded_charge_codes`/`skip_dg_generation` end-to-end coverage.
- Not yet checked: the sibling "AUS NEA to AUWC TIER 1" reference pair
  (folders 35/36).

## 3.23 Charge-code names: imported the TAD tool's 586-code lookup table

At the user's request, went through `reference/Tool_for_TAD_DirectExport.bas`
and the live `reference/Tool for TAD.xlsm` to understand the team's
in-house VBA export engine before eventually building the TAD FILING
parser (full writeup: see the `project_tad_vba_tool_analysis` memory).
One concrete outcome landed in code now, independent of the TAD build
itself:

- The workbook's `Sheet1` turned out to be a 586-row (575 unique)
  charge-code -> full-name lookup table - the VBA tool's own equivalent
  of `schema/charge_codes.py::CHARGE_CODE_NAMES`, and far more complete
  than this project's hand-built 13-entry version. User approved
  importing it wholesale as `_TAD_TOOL_CHARGE_CODE_NAMES`, merged so the
  13 already ground-truth-confirmed entries always win on overlap (the
  one real conflict, `HEA`, kept this project's confirmed "HEAVY WEIGHT
  SURCHARGE" over the TAD sheet's "HEAVY SURCHARGE"). `CHARGE_CODE_NAMES`
  now has 575 entries total.
- `INDIVIDUAL_CHARGE_CODES` (the narrower whitelist gating which codes
  auto-qualify for a raw-text-parsed "Includes" line) was deliberately
  **NOT** expanded to match - unchanged at 13. A code having a name
  available now doesn't mean it should auto-generate a CMDT NOTE child
  row for any lane; that whitelist still only grows per-lane, per-
  ground-truth-confirmation, same rule as before.
- Confirmed via full test suite (108/108 unchanged) that this is a
  pure addition with no regressions.

## 5. Files touched this session (everything above is new)

Nothing pre-existed before this session — the whole `mrg2opus/` package,
`tests/`, `data/`, `reference/`, `README.md`, `.gitignore`, and
`requirements.txt` were built from scratch. If diffing against an earlier
snapshot, treat all of it as new.

The LAWC parser was built and debugged in a later part of this same
session (after this file's first draft, which had LAWC marked "not
started" — that status is now stale, see §3 above). New/changed for LAWC:
`parsers/lawc.py`, `config/container_maps/lawc.yaml`,
`tests/test_parsers_lawc.py`, `cli.py` (import line), and one shared-code
change: `parsers/common/header_grid.py::flatten_pod_header()` gained an
optional `fallback_container_cycle` parameter (default `None`, backward
compatible) for the one real data-entry gap found in LAWC's raw sheet
(a destination with all 3 container-label cells blank).

## 6. Known limitations (won't surprise you if you re-derive them, but save the time)

**Cross-lane CMDT NOTE contradictions** (why `cmdt_notes.py` has so many
knobs): whether the "inclusive of X and Y" text is alphabetized or
preserves input order differs by lane (CSE: input order; SAF/EAF/LAEC:
alphabetized — could not be distinguished for SAF/EAF since their input
happened to already be alphabetical). Charge-code full names can differ
per lane for the same code (`"HEA"` = "HEAVY SURCHARGE" in LAEC,
"HEAVY WEIGHT SURCHARGE" in EAF). Child-row order is sometimes NOT
alphabetical even when the text is. Whether `route_seq`/`cmdt_seq`/
`commodity_note` populate the grouped sheet, the exploded PORT-PORT sheet,
neither, or both, varies by lane and even between EAF's own two sub-lanes
in the same workbook — treat each lane's convention as independently
verified, never assumed.

**Per-lane documented gaps** (each has a code comment at its source and a
test exclusion — search each lane's test file for the exact fields):
- SAF/EAF: a handful of raw-spelling location typos closed via
  `location_bank/known_aliases.py`; EAF's PORT-PORT `route_seq`/`cmdt_seq`/
  `commodity_note` inconsistency between TZDAR and KEMBA (left blank).
- CSE: Prefix O/F `PAMIT` `rate_20` (raw column looks stale for 66/69
  origins); a handful of inland-China ARBS place names where raw text and
  the Location Bank each get some codes right and some wrong; one CMDT
  NOTE `pol="BDCGP"` field tied to an unclear Chittagong override.
- LAEC: same `pol="BDCGP"` gap; `route_seq` on the grouped RATES sheet
  resets at Non-ISC/ISC boundaries with no derivable rule (excluded from
  test); OPUS ARBS is generated (user-requested, reusing CSE's verified
  logic) but this lane's ground-truth sample has no ARBS sheet to check it
  against; OPUS SPECIAL NOTE intentionally NOT generated even though the
  source PSA table exists in "IMO charge" — unlike ARBS, the user hasn't
  confirmed this is wanted.
- LAWC: same `pol` gap category (SEA group's CSS/THL/DOC/CDD child rows
  carry `pol="LKCMB"`); `cmdt_seq`/`route_seq` externally-assigned on
  RATES; `commodity_note` on RATES and RATES PORT-PORT is excluded from
  the golden-diff — it's derived correctly (matches CMDT NOTE's own text
  1:1) but this ground truth sample's own RATES-sheet column is internally
  desynced from its `commodity_group_code`, and PORT-PORT's column holds a
  bare integer instead of text at all. Neither is a parser bug; see
  `tests/test_parsers_lawc.py`'s comments for the full reasoning.
  **New (§3.14):** against the newer `reference/2_OPUS/15_LAWC FAK` real
  file, ISC/SEA's HNSLO destination rows aren't fully detected (route-note
  count short: MAR 190 vs 196, MX2 entirely 0 vs 88) — real, confirmed,
  unrelated to route notes, deliberately not chased down this session;
  see `tests/test_lawc_route_notes_reference.py`'s docstring.
- Project-wide: Freetime (POD free-days + notes) is out of scope — no
  OPUS sheet in any of the 5 samples demonstrates a target format for it.

**Formerly a real bug, now moot** — `commodity_bank/` used to mine a flat
`{code: description}` JSON across all 5 samples without lane-namespacing
(commodity group codes like `G0003` are lane-local, not globally unique,
so later-alphabetical lanes would silently clobber earlier ones' entries).
Nothing ever read from that file, so it never actually caused a wrong
description anywhere - and the whole package was removed in the session
that added user-owned commodity codes/descriptions (§3.6), which made the
bug moot rather than fixed.

## 7. Pending tasks, roughly in order

All 5 sample lanes (SAF, EAF, CSE, LAEC, LAWC) and Phase 2 (Streamlit
wizard, §3.5, plus the commodity-code/type changes in §3.6) are now done.
Remaining:

1. **Location Bank match-confidence tracking** — needed for the plan's
   Step 3 "review low-confidence matches" flow (§3.5's first "deliberately
   not built" item). Requires changing every parser's `to_opus_rows()` to
   return which raw tokens it resolved and at what confidence, not just
   silently drop unresolvable rows - a parser contract change, plan before
   starting.
2. **Phase 3 (validation & presets)** — not started. `audit/` package
   stub exists but has no rule functions yet; the UI has no Step 3.5 slot
   to show them in until this exists. Preset save/load (`presets/store.py`)
   IS done as part of Phase 2, ahead of the original phase split.
3. **UN/LOCODE ingestion** — `reference/unlocode/` exists but is empty;
   `location_bank/bootstrap_unlocode.py` was never built. Current Location
   Bank is 100% sample-mined + a handful of curated aliases. Would close
   remaining fuzzy-match/description gaps (e.g. the CSE ARBS place-name
   issues in §6) if the user ever supplies the CSV.
4. Not asked for yet, but worth flagging if it comes up: Freetime
   extraction has literally zero worked examples across all 5 lanes' OPUS
   sheets — if the user wants this, it needs a real target schema from
   them first, not a guess.
5. The 4 unpaired raw-only samples in `MRGs RAW SAMPLES/` (no ground
   truth) have never been run through the classifier/parsers end-to-end —
   worth a smoke test (upload through the new UI, or `cli.py parse`) once
   there's a reason to trust `classify()` against genuinely unseen input,
   not just the 5 paired samples it was tuned against.

## 8. How to verify you're starting from a working state

```bash
cd "C:\Users\romsae-desktop\claude\PROCESS INNOVATION HACKATHON"
./.venv/Scripts/python.exe -m pytest tests/ -v
```

Should show 108 passed. If not, something regressed since this note was
written — bisect before building on top of it.

To check the UI itself is still working end-to-end, run
`./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py` and walk
through all 4 steps with any sample workbook.
