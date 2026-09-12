"""The settings editor is rendered on two screens now (Convert and
Compare), so what it keeps in session_state has to stay separated per
screen - a shared key would have one screen silently editing the other's
draft."""
from __future__ import annotations

from unittest import mock

import streamlit as st

from mrg2opus.presets.models import MappingProfile
from mrg2opus.ui import filing_settings
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


def test_loading_a_preset_stages_it_and_clears_the_widgets():
    """Both halves matter. Streamlit renders a keyed widget from its own
    stored value, not from the `value=` it is passed - so replacing the
    profile alone changes nothing on screen, which is what Convert's old
    Load preset button did. Clearing the keys is what makes the loaded
    settings actually appear."""
    st.session_state["convert_generate_dg"] = False
    preset = MappingProfile(name="Asia-Europe", commodity_code_overrides={"FAK": "LWE01"})

    with mock.patch.object(filing_settings, "load_preset", return_value=preset):
        filing_settings._load_preset_into("convert", "Asia-Europe")

    assert "convert_generate_dg" not in st.session_state
    assert st.session_state["convert" + filing_settings._PENDING_PRESET] is preset


def test_the_staged_preset_survives_the_widget_clearing_that_comes_with_it():
    """It is stashed under a key that does NOT start with "<prefix>_",
    because that is exactly the shape reset_filing_settings() wipes."""
    key = "convert" + filing_settings._PENDING_PRESET

    st.session_state[key] = MappingProfile(name="kept")
    filing_settings.reset_filing_settings("convert")

    assert st.session_state[key].name == "kept"


def test_one_screen_s_preset_does_not_stage_onto_the_other():
    # session_state is process-wide in bare mode, so start from nothing
    # rather than from whatever an earlier test left staged.
    for prefix in ("convert", "compare"):
        st.session_state.pop(prefix + filing_settings._PENDING_PRESET, None)

    with mock.patch.object(filing_settings, "load_preset", return_value=MappingProfile(name="x")):
        filing_settings._load_preset_into("compare", "x")

    assert "compare" + filing_settings._PENDING_PRESET in st.session_state
    assert "convert" + filing_settings._PENDING_PRESET not in st.session_state
