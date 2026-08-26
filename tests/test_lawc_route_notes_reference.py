"""Route-note verification against reference/2_OPUS's real LAWC ground
truth (delivered 2026-08-26 - see project-opus-note-sheet-taxonomy
memory), kept separate from test_parsers_lawc.py's older bundled-sample
tests because that sample predates route notes entirely for two cases
(see that file's RATES_IGNORE_FIELDS route_note comment). One raw MRG
input file, one separately-filed real OPUS output file - matching how
these were actually delivered (unlike the bundled sample, which combines
both in one workbook).
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

import openpyxl
import pytest

from mrg2opus.audit.compare import diff_by_key, rates_row_key, read_rates_sheet, read_route_note_sheet
from mrg2opus.parsers.lawc import LAWCParser
from mrg2opus.presets.models import MappingProfile

REFERENCE_DIR = Path(__file__).resolve().parents[1] / "reference"
RAW_PATH = (
    REFERENCE_DIR / "1_MRGs" / "15_LAWC FAK"
    / "20260812_MRG guideline template China_HKG_SIN_TWN_KR (15-21 Aug) and SEA ISC (15-31 Aug)_FAK (1).xlsx"
)
OPUS_PATH = REFERENCE_DIR / "2_OPUS" / "15_LAWC FAK" / "LWE ( 20260815 - 20260821 ).xlsx"

pytestmark = pytest.mark.skipif(
    not RAW_PATH.exists() or not OPUS_PATH.exists(),
    reason="reference/ ground-truth files not present in this checkout",
)

# Reefer/NOR (prefix "R") rows are excluded from both checks below: this
# reference file's own commodity_group_code for NOR is G0004, but that
# code is a user-customizable override (confirmed - see
# project-opus-note-sheet-taxonomy memory), not a derivable default, so a
# fresh parse with no override applied can't be expected to reproduce
# whatever code this particular filing batch was actually override'd to.
# Separately, NOR's DR rows don't get a DG-duplicate at all yet (a real,
# but out-of-scope-for-this-task gap - see that same memory) - which is
# the other reason its route notes (REEFER DRY AS DANGEROUS combined with
# AX3) can't be reproduced here. HNSLO/OOG/the main SEA grid's own DG rows
# are unaffected by either gap and are exactly what this test verifies.
_NOR_PREFIX = "R"


def _run_lawc_fak():
    wb = openpyxl.load_workbook(RAW_PATH, data_only=True)
    return LAWCParser().run(wb, MappingProfile())


def test_lawc_route_note_text_matches_reference_rates_sheet():
    """Every non-NOR RATES row's route_note text (HNSLO's MAR/MX2, OOG's
    KCI/OH/OWOH/OW cases including the blank-for-plain-in-gauge fix, and
    the main SEA grid's own G0004 DG rows' REEFER DRY AS DANGEROUS note)
    matches the real filed RATES sheet, keyed the same way
    test_parsers_lawc.py's own RATES tests key rows. Only asserts on the
    matched intersection (not missing/extra) - this reference file has
    known, separately-tracked gaps unrelated to route notes (see the
    _NOR_PREFIX comment, and the general "LAWC Tier 1/FAK real-file
    fidelity" follow-up noted in project-opus-note-sheet-taxonomy)."""
    row_set = _run_lawc_fak()
    generated = [r.model_dump() for r in row_set.rates if r.prefix != _NOR_PREFIX]

    ref_wb = openpyxl.load_workbook(OPUS_PATH, data_only=True, read_only=True)
    expected = [r for r in read_rates_sheet(ref_wb, "RATES") if r.get("prefix") != _NOR_PREFIX]

    result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=["route_note"])
    assert result.matched > 0, "expected at least some matched rows to actually verify route_note against"
    assert not result.field_mismatches, (
        f"{len(result.field_mismatches)} route_note mismatches, e.g. {result.field_mismatches[:10]}"
    )


def test_lawc_route_notes_rn_sheet_content_matches_reference():
    """The derived RN sheet's note-text counts, for every category this
    task actually built (OOG's OH/OW/OWOH/KCI combinations, and NOR's own
    DG-duplicate rows filed as D/DG with "REEFER DRY AS DANGEROUS" per the
    user's 2026-08-26 clarification - see SEA_DG_ROUTE_NOTE's docstring),
    exactly match the real RN sheet - ignoring header_seq/route_seq/
    note_seq (writer-assigned placeholders, not reproductions of any real
    OPUS-assigned number, see _derive_route_notes' docstring).

    HNSLO's MAR/MX2 counts are deliberately NOT checked here: this
    specific reference file's ISC/SEA sheets have a real, pre-existing
    HNSLO-detection gap unrelated to route notes (confirmed: MAR is short
    190 vs 196, MX2 is entirely 0 vs 88, while every category actually
    built in this task matches exactly) - the same "LAWC FAK/Tier 1
    real-file fidelity" gap already flagged as a separate follow-up in
    project-opus-note-sheet-taxonomy, not something this task touched."""
    row_set = _run_lawc_fak()
    generated = Counter(r.contents for r in row_set.route_notes)

    ref_wb = openpyxl.load_workbook(OPUS_PATH, data_only=True, read_only=True)
    expected = Counter(r["contents"] for r in read_route_note_sheet(ref_wb, "RN") if r["contents"])

    built_this_task = {k for k in set(generated) | set(expected) if "REEFER" in k or "KCI" in k or k in ("OH", "OWOH", "OW")}
    assert built_this_task, "expected at least some of the categories this task built to actually be present"
    mismatches = {k: (generated.get(k, 0), expected.get(k, 0)) for k in built_this_task if generated.get(k, 0) != expected.get(k, 0)}
    assert not mismatches, f"generated vs expected counts differ: {mismatches}"
