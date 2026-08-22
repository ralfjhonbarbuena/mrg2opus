"""Session-state helpers for the Streamlit wizard. Kept separate from
app.py/steps/* so each step module can read/write state without importing
streamlit's global session_state directly everywhere.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import streamlit as st

from mrg2opus.presets.models import MappingProfile


@dataclass
class WizardState:
    step: int = 1
    # Names of every uploaded file, in upload order - some lanes (e.g. CSE)
    # ship as more than one real-world file (a main file plus a separate
    # one for a specific service, like CSE's "...for VELAG and VEPBL"
    # Venezuela file). `workbook` below is always the SINGLE merged result
    # (see excel_io/merge.py) regardless of how many files were uploaded -
    # every downstream step only ever deals with one workbook.
    upload_names: list[str] = field(default_factory=list)
    # Content fingerprint (names + sha256 of each file's bytes) of the
    # current upload. Everything downstream is cached against THIS, not
    # against upload_names alone: re-uploading a file whose name is
    # unchanged but whose contents were edited must re-parse, otherwise the
    # wizard silently keeps serving the previous file's results.
    upload_key: str | None = None
    workbook: Any = None  # openpyxl Workbook - not serialized, lives in-memory for the session
    classification_results: list[Any] = field(default_factory=list)  # list[ClassificationResult], best first
    selected_lane_id: str | None = None
    profile: MappingProfile = field(default_factory=MappingProfile)
    row_sets: dict[str, Any] | None = None  # dict[suffix, OpusRowSet], set after Step 2/3 runs
    # (structural_code, description) pairs from the FIRST parse of this
    # upload, before any overrides are applied - captured once and held
    # stable so Step 3's editor always keys off the parser's own default
    # codes, never off a previously-applied override (row_sets after a
    # re-run reflects the OVERRIDDEN codes, which would make a second
    # round of edits key against the wrong dict).
    default_commodity_groups: list[tuple[str, str]] = field(default_factory=list)
    output_filename: str | None = None
    output_bytes: bytes | None = None


def get_state() -> WizardState:
    if "wizard" not in st.session_state:
        st.session_state.wizard = WizardState()
    return st.session_state.wizard


def reset_state() -> None:
    st.session_state.wizard = WizardState()


def goto(step: int) -> None:
    get_state().step = step
    st.rerun()
