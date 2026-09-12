from __future__ import annotations

import streamlit as st

from mrg2opus.parsers.registry import get_profile
from mrg2opus.ui.filing_settings import render_filing_settings
from mrg2opus.ui.parsing import run_parser
from mrg2opus.ui.sheets import output_sheets
from mrg2opus.ui.state import WizardState


def render(state: WizardState) -> None:
    st.subheader("Customize")

    if not state.row_sets:
        st.warning("Nothing parsed yet - go back to Preview.")
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
        return

    # Presets live with the settings they save (ui/filing_settings.py),
    # so Compare has them too and neither screen's copy can drift.
    #
    # The old copy here had both halves of this wrong, invisibly: Load
    # replaced state.profile without clearing the widgets, which render
    # from their own stored values, so nothing on screen changed; and
    # Save wrote state.profile, the LAST APPLIED settings, rather than
    # the ones being looked at.
    #
    # The settings themselves live there because
    # Compare needs the identical editor - the auditor drafts the filing
    # independently rather than inheriting this one.
    pending_profile = render_filing_settings(
        state.profile, state.default_commodity_groups, state.commodity_groups_by_scope,
        state.dg_twin_groups, state.reefer_nor_groups, list(state.row_sets),
        state.selected_lane_id, key_prefix="convert",
    )

    st.markdown("#### Skip output sheets")
    st.caption("Named exactly as they'll appear in the exported workbook. Every sheet the export would contain is listed.")
    sheets = output_sheets(state.row_sets, get_profile(state.selected_lane_id).parser_cls)
    skip_choices: dict[str, bool] = {}
    if sheets:
        cols = st.columns(min(3, len(sheets)))
        for i, sheet in enumerate(sheets):
            with cols[i % len(cols)]:
                skip_choices[sheet.name] = st.checkbox(
                    f"{sheet.name}  ({sheet.rows:,})",
                    value=state.profile.skip_output_sheets.get(sheet.name, False),
                    key=f"skip_{sheet.scope}_{sheet.name}",
                )
    else:
        st.caption("No output sheets to skip.")

    col_back, col_next = st.columns(2)
    with col_back:
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
    with col_next:
        if st.button("Apply & Continue to Export →", type="primary"):
            # The settings above are only adopted here, on Apply: the
            # component builds a profile from the screen every rerun, but
            # a half-made edit shouldn't re-parse the workbook behind the
            # user. Skipped sheets are Convert's own - they're about the
            # exported workbook, not about how the filing is built.
            skip_output_sheets = {name: skip for name, skip in skip_choices.items() if skip}
            state.profile = pending_profile.model_copy(
                update={"skip_output_sheets": skip_output_sheets}
            )

            parser_cls = get_profile(state.selected_lane_id).parser_cls
            parser = parser_cls()
            with st.spinner("Re-running with overrides..."):
                state.row_sets = run_parser(parser, state.workbook, state.profile)
            state.output_bytes = None  # invalidate any previously-built export
            state.step = 4
            st.rerun()
