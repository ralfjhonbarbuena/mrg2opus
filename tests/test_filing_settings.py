"""The settings editor is rendered on two screens now (Convert and
Compare), so what it keeps in session_state has to stay separated per
screen - a shared key would have one screen silently editing the other's
draft."""
from __future__ import annotations

import streamlit as st

from mrg2opus.presets.models import MappingProfile, ScopeOverrides
from mrg2opus.presets.store import export_profile, import_profile, preset_filename
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


class _Upload:
    """Stands in for Streamlit's UploadedFile - name plus bytes."""

    def __init__(self, name: str, data: bytes) -> None:
        self.name, self._data = name, data

    def getvalue(self) -> bytes:
        return self._data


def _import(prefix: str, upload) -> None:
    key = f"{prefix}_preset_import"
    st.session_state[key] = upload
    filing_settings._import_preset(prefix, key)


def test_importing_a_settings_file_stages_it_and_clears_the_widgets():
    """Both halves matter. Streamlit renders a keyed widget from its own
    stored value, not from the `value=` it is passed - so replacing the
    profile alone changes nothing on screen, which is what the old Load
    preset button did. Clearing the keys is what makes the imported
    settings actually appear."""
    st.session_state["convert_generate_dg"] = False
    exported = export_profile(MappingProfile(name="Asia-Europe", commodity_code_overrides={"FAK": "LWE01"}))

    _import("convert", _Upload("Asia-Europe.json", exported.encode()))

    staged = st.session_state["convert" + filing_settings._PENDING_PRESET]
    assert staged.commodity_code_overrides == {"FAK": "LWE01"}
    assert "convert_generate_dg" not in st.session_state


def test_the_staged_preset_survives_the_widget_clearing_that_comes_with_it():
    """It is stashed under a key that does NOT start with "<prefix>_",
    because that is exactly the shape reset_filing_settings() wipes."""
    key = "convert" + filing_settings._PENDING_PRESET

    st.session_state[key] = MappingProfile(name="kept")
    filing_settings.reset_filing_settings("convert")

    assert st.session_state[key].name == "kept"


def test_a_file_that_is_not_settings_says_so_and_changes_nothing():
    """Handing back an empty profile would file the lane's own defaults
    under the user's belief that their settings had loaded."""
    for prefix in ("convert",):
        st.session_state.pop(prefix + filing_settings._PENDING_PRESET, None)

    _import("convert", _Upload("holiday-photo.json", b"{not json at all"))

    assert "convert" + filing_settings._PENDING_PRESET not in st.session_state
    assert "holiday-photo.json" in st.session_state["convert" + filing_settings._IMPORT_ERROR]


def test_one_screen_s_import_does_not_stage_onto_the_other():
    for prefix in ("convert", "compare"):
        st.session_state.pop(prefix + filing_settings._PENDING_PRESET, None)

    _import("compare", _Upload("x.json", export_profile(MappingProfile(name="x")).encode()))

    assert "compare" + filing_settings._PENDING_PRESET in st.session_state
    assert "convert" + filing_settings._PENDING_PRESET not in st.session_state


def test_an_exported_file_carries_every_setting_back():
    """Including the per-scope overrides, which are a nested model - the
    part most likely to be lost by a hand-rolled serializer."""
    profile = MappingProfile(
        name="round-trip",
        commodity_code_overrides={"FAK": "LWE01"},
        skip_dg_generation={"OOG": True},
        excluded_charge_codes=["BAF"],
        include_vertical_rates=False,
        by_scope={"AMW": ScopeOverrides(commodity_code_overrides={"FAK": "G0011"})},
    )

    back = import_profile(export_profile(profile))

    assert back.commodity_code_overrides == {"FAK": "LWE01"}
    assert back.skip_dg_generation == {"OOG": True}
    assert back.excluded_charge_codes == ["BAF"]
    assert back.include_vertical_rates is False
    assert back.by_scope["AMW"].commodity_code_overrides == {"FAK": "G0011"}


def test_a_settings_file_is_named_after_the_settings():
    assert preset_filename("LAWC Tier 1") == "LAWC Tier 1.json"
    assert preset_filename("Weird/Name:*?") == "WeirdName.json"
    assert preset_filename("///") == "settings.json"


def _staged(prefix):
    return st.session_state.get(prefix + filing_settings._PENDING_PRESET)


def _clear(prefix):
    for suffix in (filing_settings._PENDING_PRESET, filing_settings._IMPORT_ERROR,
                   filing_settings._IMPORT_MISMATCH):
        st.session_state.pop(prefix + suffix, None)


def _import_for(prefix, upload, lane_id):
    key = f"{prefix}_preset_import"
    st.session_state[key] = upload
    filing_settings._import_preset(prefix, key, lane_id)


def _file_for(lane_id):
    return _Upload(
        f"{lane_id}.json",
        export_profile(MappingProfile(name=lane_id, lane_id=lane_id,
                                      commodity_code_overrides={"FAK": "WRONG"})).encode(),
    )


def test_a_preset_from_another_lane_is_refused():
    """Not the harmless case it looks like: every commodity setting is
    keyed by a group's default description, and "FAK" is a group in EAF
    and in all three TAD lanes - so the wrong lane's overrides don't fail
    to apply, they rename and recode a group that shares a name."""
    _clear("convert")

    _import_for("convert", _file_for("TAD-OEW-OMW"), "EAF")

    assert _staged("convert") is None
    problem = st.session_state["convert" + filing_settings._IMPORT_ERROR]
    assert "TAD-OEW-OMW" in problem and "EAF" in problem


def test_the_refused_preset_is_kept_for_import_anyway():
    """Some of what a preset carries - the RFA dates, the excluded charge
    codes - is the same whatever the lane, so the file is held rather
    than thrown away."""
    _clear("convert")
    _import_for("convert", _file_for("TAD-OEW-OMW"), "EAF")

    filing_settings._import_anyway("convert")

    assert _staged("convert").commodity_code_overrides == {"FAK": "WRONG"}
    assert "convert" + filing_settings._IMPORT_MISMATCH not in st.session_state


def test_a_preset_for_this_lane_imports_without_comment():
    _clear("convert")

    _import_for("convert", _file_for("EAF"), "EAF")

    assert _staged("convert").lane_id == "EAF"
    assert "convert" + filing_settings._IMPORT_ERROR not in st.session_state


def test_a_preset_with_no_lane_still_imports():
    """Files written before the stamp existed, and any built with no lane
    picked. There is nothing to check them against, so they go through."""
    _clear("convert")
    unstamped = _Upload("old.json", export_profile(MappingProfile(name="old")).encode())

    _import_for("convert", unstamped, "EAF")

    assert _staged("convert") is not None
    assert "convert" + filing_settings._IMPORT_ERROR not in st.session_state


def test_a_bad_file_does_not_leave_a_stale_mismatch_behind():
    """Otherwise "Import anyway" would still be offered, and would stage
    the file from the previous attempt."""
    _clear("convert")
    _import_for("convert", _file_for("TAD-OEW-OMW"), "EAF")

    _import_for("convert", _Upload("junk.json", b"{not json"), "EAF")

    assert "convert" + filing_settings._IMPORT_MISMATCH not in st.session_state
