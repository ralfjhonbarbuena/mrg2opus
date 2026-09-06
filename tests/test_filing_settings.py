"""The settings editor is rendered on two screens now (Convert and
Compare), so what it keeps in session_state has to stay separated per
screen - a shared key would have one screen silently editing the other's
draft."""
from __future__ import annotations

import streamlit as st

from mrg2opus.ui.filing_settings import _editor_key, _refresh_editor, reset_filing_settings


def test_the_two_screens_get_different_editor_keys():
    assert _editor_key("convert") != _editor_key("compare")


def test_refreshing_one_screen_leaves_the_other_alone():
    convert_before = _editor_key("convert")
    compare_before = _editor_key("compare")
    _refresh_editor("compare")
    assert _editor_key("compare") != compare_before
    assert _editor_key("convert") == convert_before


def test_reset_clears_only_its_own_screens_widgets():
    st.session_state["compare_generate_dg"] = True
    st.session_state["convert_generate_dg"] = True
    editor_before = _editor_key("compare")

    reset_filing_settings("compare")

    assert "compare_generate_dg" not in st.session_state
    assert st.session_state["convert_generate_dg"] is True
    # The grid keeps its edits until its KEY changes, so a reset that
    # only dropped the other widgets would leave the old lane's commodity
    # rows on screen.
    assert _editor_key("compare") != editor_before
