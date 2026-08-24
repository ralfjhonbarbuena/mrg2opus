from __future__ import annotations

import streamlit as st

from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.parsers.registry import all_profiles
from mrg2opus.presets.models import MappingProfile
from mrg2opus.ui.errors import show_error
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify
from mrg2opus.ui.state import WizardState


def render(state: WizardState) -> None:
    st.subheader("Upload raw MRG Excel file(s)")
    st.caption(
        "Some lanes ship as more than one real-world file - e.g. CSE's main file plus a separate "
        "\"...for VELAG and VEPBL\" file. Upload every file for one filing together; they're merged "
        "into a single workbook before classification."
    )

    uploaded = st.file_uploader("MRG rate sheet(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    if not uploaded:
        st.info("Upload one or more files to classify them and continue.")
        return

    names = [f.name for f in uploaded]
    payloads = [f.getvalue() for f in uploaded]
    upload_key = fingerprint_uploads(names, payloads)

    if upload_key != state.upload_key:
        # New file set (or edited contents) - reset anything downstream so a
        # stale row_sets/profile from a previous upload can't leak into this
        # one. Resetting `profile` here also matters for Step 2's
        # default_commodity_groups snapshot - it must be taken with no
        # overrides carried over from a prior upload, or it wouldn't
        # reflect this file set's true defaults.
        state.upload_names = names
        state.upload_key = upload_key
        state.workbook = None
        state.classification_results = []
        state.selected_lane_id = None
        state.profile = MappingProfile()
        state.row_sets = None
        state.default_commodity_groups = []
        state.output_bytes = None

    if state.workbook is None:
        try:
            state.workbook, state.classification_results = load_and_classify(payloads)
        except DuplicateSheetError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surfaced directly to the user, not swallowed
            show_error(
                "Couldn't open one of these as an Excel workbook. It may not be a valid "
                ".xlsx file, or the file could be corrupted.",
                exc,
            )
            return

    label = state.upload_names[0] if len(state.upload_names) == 1 else f"{len(state.upload_names)} files"
    st.success(f"Loaded **{label}** — sheets: {', '.join(state.workbook.sheetnames)}")

    results = state.classification_results
    best = results[0] if results else None

    st.markdown("#### Classification")
    if best is None:
        st.error("No lane parsers are registered - nothing to classify against.")
        return

    below_threshold = best.confidence < best.profile.min_confidence
    if below_threshold:
        st.warning(
            f"Best match is **{best.profile.lane_id}** at {best.confidence:.0%} confidence, below the "
            f"{best.profile.min_confidence:.0%} threshold. Pick the correct lane manually below."
        )
    else:
        st.metric("Best match", best.profile.lane_id, f"{best.confidence:.0%} confidence")

    with st.expander("Score breakdown (all registered lanes)", expanded=below_threshold):
        st.dataframe(
            [
                {
                    "lane": r.profile.lane_id,
                    "confidence": f"{r.confidence:.0%}",
                    "sheet_name": f"{r.breakdown['sheet_name']:.0%}",
                    "title_keywords": f"{r.breakdown['title_keywords']:.0%}",
                    "header_fingerprint": f"{r.breakdown['header_fingerprint']:.0%}",
                }
                for r in results
            ],
            hide_index=True,
            width="stretch",
        )

    lane_ids = [p.lane_id for p in all_profiles()]
    default_lane = state.selected_lane_id or best.profile.lane_id
    selected = st.selectbox(
        "Lane to use",
        options=lane_ids,
        index=lane_ids.index(default_lane) if default_lane in lane_ids else 0,
        help="Auto-selected from the best classification match; override if it's wrong.",
    )
    state.selected_lane_id = selected

    if st.button("Continue to Preview →", type="primary"):
        state.step = 2
        st.rerun()
