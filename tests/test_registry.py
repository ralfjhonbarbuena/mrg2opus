from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

# Import every lane module so their LayoutProfiles are registered.
from mrg2opus.parsers import cse, eaf, laec, lawc, saf, waf  # noqa: F401
from mrg2opus.parsers.registry import all_profiles, classify, classify_all, get_profile

REFERENCE_DIR = Path(__file__).resolve().parents[1] / "reference"

# All 5 lanes are still registered parsers even though SAF has no real
# reference/ file to classify against (see feedback-reference-folder-
# convention memory) - SAMPLES only covers the 4 lanes that do.
ALL_LANE_IDS = {"SAF", "EAF", "CSE", "LAEC", "LAWC", "WAF"}
SAMPLES = {
    "EAF": REFERENCE_DIR / "1_MRGs" / "5_EAF-TZDAR" / "Asia EAF rate guideline TZDAR 19 Aug to 25 Aug 26 (14 Aug updated).xlsx",
    "CSE": REFERENCE_DIR / "1_MRGs" / "1_CSE FAK, CSE FAK FOR VELAG AND VEPBL" / "CSE Pricing Guideline (15-21  AUG 2026 ) FAK.xlsx",
    "LAEC": REFERENCE_DIR / "1_MRGs" / "19_LAEC FAK" / "LAEC Pricing Guideline - CN (20260901-20260907) (FAK) _ IN (20260901-20260907).xlsx",
    "LAWC": (
        REFERENCE_DIR / "1_MRGs" / "15_LAWC FAK"
        / "20260812_MRG guideline template China_HKG_SIN_TWN_KR (15-21 Aug) and SEA ISC (15-31 Aug)_FAK (1).xlsx"
    ),
    "WAF": (
        REFERENCE_DIR / "1_MRGs" / "9_West Africa WAF"
        / "Asia WAF MRG rate (26 Aug 2026 - 01 Sep 2026) (18 Aug updated.) (1).xlsx"
    ),
}

pytestmark = pytest.mark.skipif(
    any(not p.exists() for p in SAMPLES.values()), reason="reference/ ground-truth files not present in this checkout"
)


def test_all_profiles_includes_every_lane():
    lane_ids = {p.lane_id for p in all_profiles()}
    assert lane_ids == ALL_LANE_IDS


def test_get_profile_round_trips_lane_id():
    for lane_id in ALL_LANE_IDS:
        assert get_profile(lane_id).lane_id == lane_id


def test_get_profile_unknown_lane_raises():
    try:
        get_profile("NOT-A-REAL-LANE")
    except KeyError:
        pass
    else:
        raise AssertionError("expected KeyError for an unregistered lane_id")


def test_classify_all_ranks_correct_lane_first_for_every_sample():
    for lane_id, path in SAMPLES.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        results = classify_all(wb)
        assert results, f"no results for {lane_id}"
        assert results[0].profile.lane_id == lane_id
        # sorted descending by confidence
        confidences = [r.confidence for r in results]
        assert confidences == sorted(confidences, reverse=True)


def test_classify_matches_classify_all_top_result():
    wb = openpyxl.load_workbook(SAMPLES["LAEC"], data_only=True)
    assert classify(wb).profile.lane_id == classify_all(wb)[0].profile.lane_id
