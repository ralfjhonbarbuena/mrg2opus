from __future__ import annotations

import openpyxl

# Import every lane module so their LayoutProfiles are registered.
from mrg2opus.parsers import cse, eaf, laec, lawc, saf  # noqa: F401
from mrg2opus.parsers.registry import all_profiles, classify, classify_all, get_profile

SAMPLES = {
    "SAF": "Sample MRGs with OPUS FORMATS/SAF.xlsx",
    "EAF": "Sample MRGs with OPUS FORMATS/EAF.xlsx",
    "CSE": "Sample MRGs with OPUS FORMATS/CSE.xlsx",
    "LAEC": "Sample MRGs with OPUS FORMATS/LAEC.xlsx",
    "LAWC": "Sample MRGs with OPUS FORMATS/LAWC.xlsx",
}


def test_all_profiles_includes_every_lane():
    lane_ids = {p.lane_id for p in all_profiles()}
    assert lane_ids == set(SAMPLES)


def test_get_profile_round_trips_lane_id():
    for lane_id in SAMPLES:
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
    wb = openpyxl.load_workbook(SAMPLES["SAF"], data_only=True)
    assert classify(wb).profile.lane_id == classify_all(wb)[0].profile.lane_id
