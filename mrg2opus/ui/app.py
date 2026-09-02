"""Phase 2 Streamlit wizard entrypoint.

    ./.venv/Scripts/python.exe -m streamlit run streamlit_app.py

Two modes, selected at the top: "Convert" (the 4-step wizard:
upload+classify -> preview -> customize -> export) and "Compare"
(standalone: upload an MRG plus a reference OPUS file, see where they
diverge - see docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md).

streamlit_app.py at the repo root is the canonical entry point, but this
module is ALSO runnable directly (`streamlit run mrg2opus/ui/app.py`) -
that's what an existing Streamlit Cloud deploy and the devcontainer are
configured to do, and it's the path a person naturally reaches for. See
the sys.path guard below for why that needs help.
"""
from __future__ import annotations

import sys
from pathlib import Path

# `streamlit run` puts the SCRIPT's own folder on sys.path - here that's
# mrg2opus/ui, which leaves the repo root off it entirely, so the very
# next import ("from mrg2opus.parsers import ...") raises
# ModuleNotFoundError. Running from the root with `python -m streamlit`
# happens to work, which is why this only ever showed up on deploy.
# Putting the root back means either entry point works, and no host has to
# be configured a particular way.
_ROOT = str(Path(__file__).resolve().parents[2])
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import streamlit as st  # noqa: E402  (must follow the sys.path guard)

# Importing the lane modules registers their LayoutProfile as a side effect -
# same requirement as cli.py.
from mrg2opus.parsers import (  # noqa: F401
    aubp, auec, auwc, cse, eaf, laec, lawc, nz1_sea, nzj, saf, tad_aew_amw, tad_oew_omw, tad_wmw_wew, waf, west_asia_multi, west_asia_waf,
)
from mrg2opus.ui import compare_page
from mrg2opus.ui.state import get_state
from mrg2opus.ui.steps import step1_upload, step2_preview, step3_customize, step4_export
from mrg2opus.ui.theme import apply_theme

STEP_LABELS = {
    1: "1. Upload & Classify",
    2: "2. Preview",
    3: "3. Customize",
    4: "4. Export",
}

STEP_RENDERERS = {
    1: step1_upload.render,
    2: step2_preview.render,
    3: step3_customize.render,
    4: step4_export.render,
}


def main() -> None:
    st.set_page_config(page_title="mrg2opus", layout="wide")
    apply_theme()  # must follow set_page_config; see ui/theme.py
    st.title("MRG → OPUS Converter")

    # The two options name themselves, so the "Mode" label is noise.
    mode = st.radio(
        "Mode", options=["Convert", "Compare"], horizontal=True, label_visibility="collapsed"
    )
    st.divider()

    if mode == "Compare":
        compare_page.render()
        return

    state = get_state()

    # Every step is a button of the same size, the current one filled.
    # It used to render the current step as markdown text among real
    # buttons, which left the row visibly uneven; a text tick is used for
    # "done" rather than an emoji, so all four labels share one typeface
    # and sit on one baseline.
    cols = st.columns(len(STEP_LABELS))
    for col, (step_num, label) in zip(cols, STEP_LABELS.items()):
        with col:
            prefix = "✓ " if step_num < state.step else ""
            # Every step's own render() already checks its prerequisites
            # (state.workbook/row_sets) and shows a "go back" prompt if
            # they're missing, so jumping directly to any step - not
            # just completed ones - degrades gracefully rather than
            # erroring.
            if st.button(
                f"{prefix}{label}",
                key=f"step_nav_{step_num}",
                width="stretch",
                type="primary" if step_num == state.step else "secondary",
            ):
                state.step = step_num
                st.rerun()
    st.divider()

    STEP_RENDERERS[state.step](state)


if __name__ == "__main__":
    main()
