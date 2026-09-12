"""MappingProfile presets as JSON, so a recurring lane's customization
doesn't need re-entering every run. Fails fast on a corrupted file
(pydantic validation error) rather than silently falling back to defaults
- a silently-dropped override on billing-relevant data is worse than a
visible crash.

Two ways in and out, one format. export_profile/import_profile pass the
settings as a file the filer keeps, which is what the app itself uses:
a preset then travels with the person - onto another machine, to the
auditor, into the repo - instead of living in a folder beside whichever
copy of the app happened to write it. save_preset/load_preset keep a
directory of them, which is the same bytes under a name; a file written
by either is readable by both.
"""
from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

from mrg2opus.presets.models import MappingProfile

DEFAULT_PRESETS_DIR = Path(__file__).resolve().parents[2] / "data" / "presets"

_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9 _-]+")


def _path_for(name: str, presets_dir: Path) -> Path:
    safe = _SAFE_NAME_RE.sub("", name).strip()
    if not safe:
        raise ValueError(f"Preset name {name!r} has no usable characters (letters/digits/space/-/_ only)")
    return presets_dir / f"{safe}.json"


def save_preset(profile: MappingProfile, presets_dir: Path | str = DEFAULT_PRESETS_DIR) -> Path:
    presets_dir = Path(presets_dir)
    presets_dir.mkdir(parents=True, exist_ok=True)
    stamped = profile.model_copy(update={"updated_at": datetime.now(timezone.utc).isoformat()})
    path = _path_for(stamped.name, presets_dir)
    path.write_text(stamped.model_dump_json(indent=2), encoding="utf-8")
    return path


def load_preset(name: str, presets_dir: Path | str = DEFAULT_PRESETS_DIR) -> MappingProfile:
    path = _path_for(name, Path(presets_dir))
    return MappingProfile.model_validate_json(path.read_text(encoding="utf-8"))


def list_presets(presets_dir: Path | str = DEFAULT_PRESETS_DIR) -> list[str]:
    presets_dir = Path(presets_dir)
    if not presets_dir.exists():
        return []
    return sorted(p.stem for p in presets_dir.glob("*.json"))


def delete_preset(name: str, presets_dir: Path | str = DEFAULT_PRESETS_DIR) -> None:
    _path_for(name, Path(presets_dir)).unlink(missing_ok=True)


def preset_filename(name: str) -> str:
    """A download name for a preset, by the same rule the directory uses,
    so a preset exported and one saved are named alike."""
    safe = _SAFE_NAME_RE.sub("", name).strip()
    return f"{safe or 'settings'}.json"


def export_profile(profile: MappingProfile) -> str:
    """The settings as a file's worth of JSON, stamped with the time -
    byte-identical to what save_preset() writes, so an exported file can
    be dropped into the presets directory and vice versa."""
    stamped = profile.model_copy(update={"updated_at": datetime.now(timezone.utc).isoformat()})
    return stamped.model_dump_json(indent=2)


def import_profile(data: str | bytes) -> MappingProfile:
    """Read settings back from an exported file.

    Raises rather than returning a default on anything malformed: silently
    handing back an empty profile would file the lane's own defaults under
    the user's belief that their settings had loaded.
    """
    if isinstance(data, bytes):
        data = data.decode("utf-8")
    return MappingProfile.model_validate_json(data)
