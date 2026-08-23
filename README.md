# mrg2opus

Converts raw carrier MRG (Master Rate Guideline) Excel rate sheets into the
structured OPUS filing format. See `mrg2opus/schema/`, `mrg2opus/parsers/`,
and `mrg2opus/location_bank/` for the core pieces; the full design is in
the plan this was built from.

## Setup

Python 3.14 is at `C:\Users\romsae-desktop\AppData\Local\Python\bin\python.exe`
(the WindowsApps `python`/`py` commands are unreliable store shims - don't
use them). A `.venv` already exists in this folder with all dependencies
installed; to recreate it from scratch:

```bash
"C:\Users\romsae-desktop\AppData\Local\Python\bin\python.exe" -m venv .venv
./.venv/Scripts/python.exe -m pip install -r requirements.txt
```

## One-time reference data bootstrap

Mines the Location Bank from the 5 ground-truth sample workbooks in
`Sample MRGs with OPUS FORMATS/`. Run once (or after adding new paired
samples):

```bash
./.venv/Scripts/python.exe -m mrg2opus.location_bank.bootstrap_from_samples
```

This creates `data/location_bank.sqlite3`. To add full UN/LOCODE coverage
(recommended - closes fuzzy-match gaps for locations that never appear
standalone in the 5 samples), download the UN/LOCODE CSV from UNECE
(https://unece.org/trade/cefact/UNLOCODE-Download) into
`reference/unlocode/`, then run `bootstrap_unlocode.py` (Phase 2/3 work -
not yet implemented).

There is no equivalent bank/registry for commodity group codes - each
parser ships a hardcoded default code+description per commodity group
(just a starting suggestion, verified against that lane's own ground
truth), but the code that actually ends up in the OPUS output is meant to
mainly come from the user via the Step 3 UI
(`MappingProfile.commodity_code_overrides`), not from any mined/bootstrapped
store. Descriptions also ship with a default but stay user-editable the
same way. See `mrg2opus/parsers/common/commodity.py`.

**Every raw sheet defaults to its own description, and merges only on
request.** A handful of raw sheets used to get folded into another
sheet's combined description by default (LAWC's "Reefer"/"LAWC NOR" into
the main dry grid's description; CSE's "NOR(PA)"/"CSE VE" into "CSE"'s,
and "NOR (MAOVLD)" into "CSE (MAOVLD)"'s; LAEC's "R5 NOR" into the Non-ISC
main group's) - each of these now defaults to its own description (its
own raw sheet name) and gets its own `OPUS CMDT NOTE` block. If the user
overrides two or more groups' descriptions to the exact same text (in
Step 3, or via `MappingProfile.commodity_description_overrides`), they
collapse back into ONE `OPUS CMDT NOTE` block, with their charge codes
unioned - see `parsers/common/commodity.py::build_notes_by_description()`.
Overriding every affected sheet's description back to the ORIGINAL
combined text exactly reconstructs that lane's ground-truth `OPUS CMDT
NOTE` sheet (verified by test, e.g.
`tests/test_parsers_lawc.py::test_lawc_cmdt_note_merges_when_descriptions_match`).

All three `MappingProfile.commodity_*_overrides` dicts (code, description,
cmdt_seq) are keyed by a group's DEFAULT description, not its code -
several groups can default to the SAME code (e.g. LAWC's main dry
grid/"Reefer"/"LAWC NOR" all default to `G0001`), so code can't serve as a
unique key once a lane has this kind of split. Default description is
always unique per group by construction, and the UI snapshots it once
(`WizardState.default_commodity_groups`) right after the first,
override-free parse, so it stays a valid lookup key across repeated
rounds of Step 3 editing even after overrides are applied.

## Running the converter

**Web UI (Phase 2, recommended):**

```bash
./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py
```

Opens a 4-step wizard: upload & classify → preview → customize (commodity
code/description/CMDT-Seq/order overrides, skip sheets, save/load named
presets) → export. See `mrg2opus/ui/app.py` and `mrg2opus/ui/steps/`.

Every row's `Type` column is always written as `"C"` - a deliberate,
uniform business rule applied to every lane regardless of what any single
sample's ground truth showed (some left it blank; that's treated as
sample noise, not a per-lane convention). See `schema/opus_rows.py`.

**Multiple input files per filing.** Step 1's upload accepts more than one
`.xlsx` at once - some lanes genuinely ship as separate real-world files
(e.g. CSE's main "Tier 1" file plus a separate "...for VELAG and VEPBL"
file covering Venezuela; the paired sample workbook already bundles these
into one file for convenience, same as it does for EAF's TZDAR/KEMBA).
Uploads are merged sheet-by-sheet into one workbook
(`excel_io/merge.py::merge_workbooks()`, preserving cell values, merged
ranges, and font/fill so strikethrough/blacked-out exclusion still works)
before classification, so the rest of the pipeline never knows more than
one file was involved. Two uploads with a sheet of the same name raise a
clear error rather than silently picking one. The CLI accepts multiple
positional file arguments the same way:
`python -m mrg2opus.cli parse file1.xlsx file2.xlsx --out out.xlsx`.

**Reordering commodity groups.** Step 3's commodity table has an **Order**
column - the underlying widget doesn't support mouse-drag row reordering,
so this is the equivalent: give groups numbers to control the sequence
their blocks appear in on the generated OPUS RATES / RATES PORT-PORT /
CMDT NOTE sheets (lower numbers first). Two groups given the same Order
value keep their relative default order; groups left unordered are
appended after every explicitly-ordered one. This is a pure reordering of
already-built rows/blocks (`parsers/common/ordering.py::reorder_row_set()`),
applied identically to what gets written to the downloaded file - not
just the in-browser preview.

**Opting out of DG (Dangerous Goods) duplicate rows.** Every lane's base
Dry (D/DR) rows normally also file an identical D/DG variant at the same
rate - a standing filing convention verified against ground truth, not
derived from any raw sheet. Step 3's commodity table has a **Skip DG**
checkbox per group: check it to stop that group's D/DG duplicates from
being generated at all (the D/DR rows are unaffected). Unchecked (the
default) keeps the existing behavior. EAF's TZDAR and KEMBA sub-lanes
share one toggle, same as every other per-group override in this table -
see `MappingProfile.skip_dg_generation`.

**If output looks stale, re-parse.** Step 2 caches its parse for the whole
session, and re-uploading the same filenames does not by itself force a
re-run. Use Step 2's **↻ Re-parse from source** button (it clears the
cached rows and any customizations, then re-reads the workbook) after
changing parser code, or if results look like a previous run's. Uploading
a file whose *contents* changed does invalidate the cache automatically -
the wizard fingerprints file bytes, not just names.

**Withdrawn locations are dropped, not just withdrawn rate cells.** The raw
MRG convention of marking a cancelled figure with strikethrough text or a
blacked-out fill (already used per-rate-cell everywhere) now also applies
to an origin's or destination's own name/code cell: if a trader struck
through or blacked out a POR/POD label to mean "this location isn't
offered," that location is dropped entirely - every row for a withdrawn
origin, every column for a withdrawn destination - instead of only
excluding whichever individual rate cells happened to carry the same
formatting. Marking just one of a location's cells (e.g. only the code,
not the name) is enough; it doesn't require every cell to be marked. See
`parsers/common/exclusion.py::location_is_excluded()` and
`parsers/common/header_grid.py::flatten_pod_header()`.

**Compare mode: check an existing OPUS file against its source MRG.** A
second top-level mode (switch at the top of the app, separate from the
4-step wizard) - upload the raw MRG file(s) plus an existing OPUS-format
Excel file (e.g. a filing someone already produced), and see a row-level
and field-level diff across all 5 OPUS sheet types the lane produces.
Choose whether to check the MRG as grouped (RATES), exploded (RATES
PORT-PORT), or both - both are always derived from the same parse, this
only controls what gets compared. See
`mrg2opus/audit/compare.py`/`mrg2opus/ui/compare_page.py`.

**CLI (Phase 1, still available for scripting):**

```bash
./.venv/Scripts/python.exe -m mrg2opus.cli parse "Sample MRGs with OPUS FORMATS/SAF.xlsx" --out out.xlsx
```

SAF, EAF (both sub-lanes: TZDAR, KEMBA), CSE, LAEC, and LAWC are all
implemented end-to-end (see `mrg2opus/parsers/saf.py`, `eaf.py`, `cse.py`,
`laec.py`, `lawc.py`).

CSE is by far the richest lane: 3 D2/D4/D5 rate grids across separate raw
sheets and commodity groups (main Caribbean/Central America service,
MAOVLD/Brazil, VE/Venezuela), 2 single-column reefer ("R5 NOR") sheets, an
in-gauge sheet producing a duplicate Prefix O/F row pair, a Yangtze inland-
origin sheet feeding OPUS ARBS, and (unlike SAF/EAF) origin and destination
codes given directly in the raw sheet rather than needing fuzzy matching -
just a Location Bank description lookup by code, plus two "regional group"
origin codes (FEBP, WPRD) that expand to a member-code list (see
`mrg2opus/config/group_codes/cse.yaml`). OPUS SPECIAL NOTE (the PSA DG
transshipment-via-Singapore surcharge table) comes from the "DG
surcharges" sheet, rows 21-24 - amounts are read from the sheet, but the
per-class row order and the validity start date are verified, hardcoded
constants (only one example exists to learn from - see the `PSA_*`
constants in `cse.py`).

EAF is the first multi-sub-lane example - it uses
`BaseMRGParser.run_multi()` / `excel_io.writer.write_opus_workbook_multi()`
to write each sub-lane's rates/notes into their own suffixed sheets
(`OPUS RATES-TZDAR`, `OPUS RATES-KEMBA`, ...).

Note: in real use, TZDAR and KEMBA are two separate MRG files (uploaded
one at a time, per Step 1 of the workflow) - `Sample MRGs with OPUS
FORMATS/EAF.xlsx` only bundles both raw sheets into one workbook as a
sample-data convenience. The parser doesn't assume both are present: it
scans whatever sub-lane sheet(s) exist in the given file, so a standalone
single-sheet upload produces exactly one correctly-suffixed sub-lane's
output (verified in `tests/test_parsers_eaf.py::test_eaf_standalone_single_lane_file`).

LAEC is CSE's structure again (direct origin/destination codes, FEBP/WPRD
group shorthand, D2/D4/D5 grid + single-column reefer sheet + in-gauge
O/F sheet + Yangtze-sheet ARBS) but doubled: everything splits into a
"Non-ISC" (Far East origins, G0015/G0010) and an "ISC" (Indian
Subcontinent origins, G0016/G0017) section on the same raw sheets, each
with its own CMDT NOTE charge-code set - reefer only exists for Non-ISC.
New pieces: the in-gauge sheet's destination header cells are themselves
"/"-joined multi-code groups (handled the same way as a grouped origin);
an "ECSA Add-On" sheet prices 3 extra destinations (Ushuaia, Zarate, La
Plata) as an existing destination's rate plus a fixed add-on. LAEC's
ground truth sample is missing an OPUS ARBS sheet even though the raw
"Yangtze ARB Add-on" sheet is present and structurally identical to CSE's
- confirmed by the user as an omission in how the sample was built, so
ARBS is generated by reusing the CSE-verified logic (now factored into
`parsers/common/yangtze_arbs.py`), but has no ground truth to check it
against for this specific lane. OPUS SPECIAL NOTE is not generated for
LAEC even though the same PSA surcharge table exists in its "IMO charge"
sheet - unlike ARBS, the user hasn't flagged this as an intentional gap
in the sample, so it's left out rather than assumed.

LAWC has four raw "Dry"-shaped grids (main China/TWN/SIN/HKG/KR, S.E.A/JPN/
SA/AU/NZ, ISC/LK/BD/AE/PK, plus a richer OOG sheet), each its own commodity
group and direct origin/destination codes (no fuzzy matching). Reefer and a
separate "NOR" sheet both file under the main commodity group as single-
column Prefix-R sheets, same pattern as CSE/LAEC's reefer sheets. OOG has 3
"/"-joined destination groups x 4 equipment-type column-pairs (in-gauge,
"OH", "OWOH", "OW"), verified to map to Prefix O+F twins (in-gauge/OH) or
Prefix F only (OWOH/OW). New pieces not seen in other lanes: a "Via"
column with either a 3-letter abbreviation or a full location code
(sometimes naming an inland transmode like "Barge via Sha" or "via CNYTN
RAIL" - the transmode word, when present, becomes `origin_transmode`); two
raw-sheet code typos (`CNXGG`->`CNTXG`, `INXIE`->`INIXE`, letter-
transpositions of real Location Bank codes, confirmed against the codes
ground truth actually resolves to); a "KCI service only" annotation on 2
OOG origins that changes those rows' `route_note` text; and San Lorenzo
(HNSLO)'s header cell spelling out a Trucking/Vessel-Service-Lane note
verbatim, applied as a per-commodity-group `route_note` + `destination_
transmode="Truck"` override. OPUS RATES PORT-PORT is also structurally
different from every other lane here: it uses a completely different
commodity-code/description namespace than OPUS RATES for the same
underlying data (`G0037`-`G0041` vs `G0001`-`G0004`, with the main group's
Dry and Reefer+NOR portions split into two separate PORT-PORT codes), and
it re-splits multi-code origin/destination groups to each port's own
individual description rather than keeping the combined ";"-joined string
verbatim like CSE/LAEC/SAF do (see `lawc.py`'s `_explode_lawc()` and
`PP_COMMODITY`).

## Web UI scope notes

The Step 3 "customize" screen's commodity-group table is built from the
codes/descriptions the parser produced on the FIRST parse of the uploaded
file (captured once as `WizardState.default_commodity_groups`, before any
overrides are applied) - not from any commodity registry, since there
isn't one; codes/descriptions are meant to come from the user. Both the
code and description columns are directly editable; edits become
`MappingProfile.commodity_code_overrides` / `commodity_description_overrides`
entries keyed by the group's DEFAULT description (not its code - see the
bootstrap section above for why), and the parser is re-run before moving
to Export. Overriding two groups to the same description merges their
`OPUS CMDT NOTE` block - see the bootstrap section above. The table's
**Order** column writes `MappingProfile.commodity_group_order` (the FINAL,
post-override description values in the row order shown, not the
defaults) and is applied via `reorder_row_set()` to RATES/CMDT NOTE
always, and to RATES PORT-PORT except on lanes where that sheet uses a
different commodity description than RATES for the same group (only
LAWC, via its PP_COMMODITY remap) - PORT-PORT keeps its default order on
that lane specifically, since there's nothing in `commodity_group_order`
to match its rows against. Two plan items are intentionally not built
yet:
the Location Bank low-confidence match review (needs every parser's
`to_opus_rows()` to report which raw tokens it resolved and at what
confidence - a parser contract change, not just a UI addition) and the
Step 3.5 Audit Gate (the underlying `audit/checks.py` rule functions are
Phase 3 work and don't exist yet).

## Tests

```bash
./.venv/Scripts/python.exe -m pytest tests/ -v
```

Tests are golden-file diffs against the OPUS sheets already present in the
paired sample workbooks - the only ground truth available. SAF and EAF
(both sub-lanes) match exactly on RATES and RATES PORT-PORT row content and
on CMDT NOTE - the Location Bank bootstrap uses a deterministic elimination
pass (see `location_bank/bootstrap_from_samples.py`) to recover names for
codes that never appear as a clean standalone row, plus a small curated
alias file (`location_bank/known_aliases.py`) for verified raw-spelling
variants (e.g. "Busan" vs. mined "PUSAN", a "HoChiMing" typo in EAF KEMBA).

Remaining known limitations, not yet addressed:

- Freetime (POD free-days + notes): no OPUS sheet in any sample demonstrates
  a target format for this, so it's intentionally not extracted yet.
- EAF's `route_seq`/`cmdt_seq`/`commodity_note` on PORT-PORT rows: TZDAR's
  ground truth cross-references the CMDT NOTE on every row, but KEMBA's
  (same workbook) leaves them blank - the two sub-lanes contradict each
  other, so this isn't generated rather than guessing which is the rule.
- CSE's Prefix O/F PAMIT `rate_20` (in-gauge 20' rate): the raw sheet's own
  column reads a flat, un-varying value for 66 of 69 origins while every
  other field (rate_40, every other destination) varies correctly per
  origin - looks like a stale/un-updated raw column, not derivable from
  anything else in the file (see `tests/test_parsers_cse.py`).
- A handful of inland Chinese ARBS place-name descriptions (documented in
  `tests/test_parsers_cse.py::ARBS_DESCRIPTION_GAP_CODES`) need real
  UN/LOCODE data to resolve correctly - the raw text and the mined
  Location Bank each get some of them right and some wrong, with no clean
  rule to arbitrate between them from this file alone.
- There is no shared commodity code registry (there was briefly a
  `commodity_bank/` mining package, removed - it wasn't lane-namespaced,
  since commodity group codes like "G0003" are lane-local, not globally
  unique, and nothing actually read from it anyway). Codes/descriptions
  are meant to come from the user (Step 3 UI), starting from each parser's
  own verified per-lane defaults - see `parsers/common/commodity.py`.
- LAEC's `route_seq` on the grouped OPUS RATES sheet: a real per-row
  sequence in the ground truth, but the numbering resets at Non-ISC/ISC
  section boundaries in a way not derivable with confidence in the time
  available (see `tests/test_parsers_laec.py`). `commodity_note` (the
  CMDT NOTE Contents text copied into every grouped row) IS implemented,
  since that part is a clean per-commodity-group lookup.
- LAEC's OPUS CMDT NOTE `pol` field: the same unresolved gap found in CSE
  - one specific THL child row carries `pol="BDCGP"` (Chittagong), tied to
  a Chittagong-specific override in the raw footnotes but not derivable
  with confidence from that text alone.
- LAWC's OPUS CMDT NOTE `pol` field: same category of gap as LAEC's above -
  the SEA group's CSS/THL/DOC/CDD child rows carry `pol="LKCMB"` (Colombo)
  in the ground truth, a single observed instance not confidently
  generalizable from raw footnote text.
- LAWC's `cmdt_seq`/`route_seq` on OPUS RATES: externally-assigned running
  numbers, same category as LAEC's gap above.
- LAWC's `commodity_note` on OPUS RATES and OPUS RATES PORT-PORT: derived
  correctly per-group (matches OPUS CMDT NOTE's own boilerplate text 1:1)
  but this ground truth sample's own RATES-sheet `commodity_note` column is
  internally desynced from its `commodity_group_code` (e.g. group G0003's
  note text lists a different group's charge codes), and PORT-PORT's
  column holds a bare integer instead of text at all - neither is
  derivable from any raw source data, so both are excluded from the
  golden-diff comparison (see `tests/test_parsers_lawc.py`).
