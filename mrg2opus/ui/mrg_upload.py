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
