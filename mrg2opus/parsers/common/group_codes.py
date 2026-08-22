"""Load "regional base-port group" expansions from a per-lane YAML config.

Some lanes (confirmed: CSE) use a short group code as shorthand for "this
rate applies to all these ports" directly in the origin column, instead of
listing every port as its own row (e.g. "FEBP" = a set of Far East base
ports). The OPUS ground truth expands these to the full member-code list.

These are curated YAML, not parsed from the raw sheet's footnote text: the
footnote text (e.g. "FEBP = current base port : Pusan/Shanghai/...") was
tried first and found unreliable against CSE's own ground truth - it lists
Korea (KRPUS) as part of FEBP even though Korea also has its own standalone
origin row and the ground truth excludes it from the FEBP expansion, and
it has typos (e.g. "Zhoaqing" for "Zhaoqing") that break clean matching.
Each entry here was instead verified directly against the ground truth
OPUS RATES sheet's actual expanded origin_code list.
"""
from __future__ import annotations

from pathlib import Path

import yaml

GROUP_CODES_DIR = Path(__file__).resolve().parents[2] / "config" / "group_codes"


def load_group_codes(lane_id: str) -> dict[str, list[str]]:
    path = GROUP_CODES_DIR / f"{lane_id.lower()}.yaml"
    if not path.exists():
        return {}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return {code: list(members) for code, members in data.items()}
