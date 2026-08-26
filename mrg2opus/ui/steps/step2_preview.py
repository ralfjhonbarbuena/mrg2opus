from __future__ import annotations

import pandas as pd
import streamlit as st

from mrg2opus.parsers.registry import get_profile
from mrg2opus.presets.models import MappingProfile
from mrg2opus.ui.commodity_utils import assign_sequential_default_codes, distinct_commodity_groups
from mrg2opus.ui.errors import show_error
from mrg2opus.ui.parsing import run_parser
from mrg2opus.ui.state import WizardState

SHEET_KEYS = ["rates", "rates_port_port", "arbs", "cmdt_notes", "special_notes"]
SHEET_LABELS = {
    "rates": "OPUS RATES",
    "rates_port_port": "OPUS RATES PORT-PORT",
    "arbs": "OPUS ARBS",
    "cmdt_notes": "OPUS CMDT NOTE",
    "special_notes": "OPUS SPECIAL NOTE",
}


def _run_parser(state: WizardState, profile: MappingProfile) -> dict:
    parser_cls = get_profile(state.selected_lane_id).parser_cls
    parser = parser_cls()
    return run_parser(parser, state.workbook, profile)


def render(state: WizardState) -> None:
    if state.workbook is None or state.selected_lane_id is None:
        # Reachable now that the step navigator (app.py) lets a user jump
        # straight here - previously this step was only ever reached via
        # Step 1's own "Continue to Preview" button, which doesn't appear
        # until a file is uploaded and classified, so this case was
        # unreachable. Without this guard, _run_parser() below would fail
        # on a None lane_id with a generic "Parsing failed" - misleading
        # when the real issue is "you haven't uploaded anything yet".
        st.warning("Nothing uploaded yet - go back to Upload & Classify.")
        if st.button("← Back to Upload"):
            state.step = 1
            st.rerun()
        return

    st.subheader(f"Preview — lane {state.selected_lane_id}")

    if state.row_sets is None:
        with st.spinner("Parsing..."):
            try:
                state.row_sets = _run_parser(state, state.profile)
            except Exception as exc:  # noqa: BLE001 - surfaced directly, this is a real parse failure
                show_error(
                    "Parsing failed - this usually means something in the file doesn't match "
                    "what this lane's parser expects.",
                    exc,
                )
                if st.button("← Back to Upload"):
                    state.step = 1
                    st.rerun()
                return
            # Snapshot the parser's own default (code, description) pairs
            # from this FIRST parse, before Step 3 applies any overrides -
            # see WizardState.default_commodity_groups.
            state.default_commodity_groups = distinct_commodity_groups(state.row_sets)

            # User-directed (2026-08-27): every distinct group gets its own
            # unique output code (G0001, G0002, ...) by default, instead of
            # silently sharing whatever structural code a lane's parser
            # happens to use internally for unrelated joins - see
            # commodity_utils.assign_sequential_default_codes' docstring.
            # Auto-seeding commodity_code_overrides (rather than changing
            # any parser's own internal code) means Step 3's existing
            # override editor, and every parser's resolve_commodity_code()
            # call, pick this up for free - re-parsing once more so
            # row_sets (this preview, and Export if untouched) reflects it
            # immediately, not just Step 3's editor. Only seeded when the
            # profile has no code overrides yet, so this never clobbers a
            # loaded preset or a re-parse after the user already customized.
            if not state.profile.commodity_code_overrides:
                state.profile = state.profile.model_copy(
                    update={"commodity_code_overrides": assign_sequential_default_codes(state.default_commodity_groups)}
                )
                state.row_sets = _run_parser(state, state.profile)

    row_sets = state.row_sets

    # Escape hatch: row_sets is cached for the whole session (parsed once,
    # above), so anything that changes the parse WITHOUT changing the
    # uploaded bytes - most often a parser code change during development -
    # would otherwise keep showing the previous run's results with no way
    # to refresh short of re-uploading. Clears the override profile too,
    # since default_commodity_groups is re-snapshotted from this parse and
    # must be taken override-free (see the snapshot comment above).
    if st.button("↻ Re-parse from source", help="Discards the cached parse and any customizations, then re-reads the uploaded workbook."):
        state.profile = MappingProfile()
        state.row_sets = None
        state.default_commodity_groups = []
        state.output_bytes = None
        st.rerun()

    st.markdown("#### Row counts")
    count_rows = []
    for suffix, row_set in row_sets.items():
        label = suffix or "(default)"
        for key in SHEET_KEYS:
            n = len(getattr(row_set, key))
            if n:
                count_rows.append({"sub-lane": label, "sheet": SHEET_LABELS[key], "rows": n})
    if count_rows:
        st.dataframe(count_rows, hide_index=True, width="stretch")
    else:
        st.warning(
            "No rows were produced at all - this usually means the wrong lane was selected, or every row "
            "failed Location Bank resolution. Go back and double-check the lane."
        )

    st.markdown("#### Data preview")
    suffix_options = list(row_sets.keys())
    suffix_labels = {s: (s or "(default)") for s in suffix_options}
    selected_suffix = (
        st.selectbox("Sub-lane", options=suffix_options, format_func=lambda s: suffix_labels[s])
        if len(suffix_options) > 1
        else suffix_options[0]
    )
    row_set = row_sets[selected_suffix]

    available_keys = [k for k in SHEET_KEYS if getattr(row_set, k)]
    if available_keys:
        selected_key = st.selectbox("Sheet", options=available_keys, format_func=lambda k: SHEET_LABELS[k])
        rows = getattr(row_set, selected_key)
        df = pd.DataFrame([r.model_dump() for r in rows[:500]])
        st.caption(f"Showing {len(df)} of {len(rows)} rows.")
        st.dataframe(df, width="stretch")
    else:
        st.caption("Nothing to preview for this sub-lane.")

    col_back, col_next = st.columns(2)
    with col_back:
        if st.button("← Back to Upload"):
            state.step = 1
            st.rerun()
    with col_next:
        if st.button("Continue to Customize →", type="primary"):
            state.step = 3
            st.rerun()
