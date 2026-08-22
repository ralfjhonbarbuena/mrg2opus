from __future__ import annotations

import pytest

from mrg2opus.presets.models import MappingProfile
from mrg2opus.presets.store import delete_preset, list_presets, load_preset, save_preset


def test_save_and_load_round_trip(tmp_path):
    profile = MappingProfile(
        name="Asia-Europe Standard",
        commodity_description_overrides={"G0001": "Custom Description"},
        commodity_sequence_overrides={"G0001": 5},
        skip_output_sheets={"OPUS ARBS": True},
    )
    save_preset(profile, presets_dir=tmp_path)

    loaded = load_preset("Asia-Europe Standard", presets_dir=tmp_path)

    assert loaded.name == profile.name
    assert loaded.commodity_description_overrides == profile.commodity_description_overrides
    assert loaded.commodity_sequence_overrides == profile.commodity_sequence_overrides
    assert loaded.skip_output_sheets == profile.skip_output_sheets
    assert loaded.updated_at is not None


def test_list_presets_empty_dir(tmp_path):
    assert list_presets(presets_dir=tmp_path / "does-not-exist") == []


def test_list_and_delete_presets(tmp_path):
    save_preset(MappingProfile(name="alpha"), presets_dir=tmp_path)
    save_preset(MappingProfile(name="beta"), presets_dir=tmp_path)

    assert list_presets(presets_dir=tmp_path) == ["alpha", "beta"]

    delete_preset("alpha", presets_dir=tmp_path)

    assert list_presets(presets_dir=tmp_path) == ["beta"]


def test_load_missing_preset_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_preset("nope", presets_dir=tmp_path)


def test_preset_name_sanitized_for_filesystem(tmp_path):
    save_preset(MappingProfile(name="Weird/Name:*?"), presets_dir=tmp_path)

    assert list_presets(presets_dir=tmp_path) == ["WeirdName"]
