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
