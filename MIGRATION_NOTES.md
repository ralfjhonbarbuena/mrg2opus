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

Full test suite: **54/54 passing** (`./.venv/Scripts/python.exe -m pytest tests/ -v`, ~2.5-3 min) - includes `test_excel_io_merge.py` (5 tests) and `test_ordering.py` (6 tests) added for §3.7's multi-file/ordering work, one CSE regression test added for §3.8's shifted-header-row fix, and 4 more (2 in `test_header_grid.py`, 2 in `test_parsers_cse.py`) added for §3.9's withdrawn-location exclusion.

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
in `tests/test_audit_compare.py`.

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

Should show 54 passed. If not, something regressed since this note was
written — bisect before building on top of it.

To check the UI itself is still working end-to-end, run
`./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py` and walk
through all 4 steps with any sample workbook.
