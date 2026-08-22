"""Per-lane container-label -> (Prefix, CGO Type, rate-column-suffix) mapping.

Raw MRG sheets label their rate columns with lane-specific container-type
tokens (D2/D4/D5/RD5, 20'/40'/40HC/45', 20/40/HCD, ...). This maps those
tokens to the fixed OPUS Prefix/CGO TYPE vocabulary and to which of the
four RatesRow rate slots (20/40/40hc/45) the value belongs in. Stored as
YAML so it's human-editable without touching parser code (the "customizable
commodity sequence, commodity group description, etc." requirement extends
naturally to this mapping too).
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

CONTAINER_MAP_DIR = Path(__file__).resolve().parents[2] / "config" / "container_maps"

VALID_SUFFIXES = {"20", "40", "40hc", "45"}


@dataclass(frozen=True)
class ContainerMap:
    prefix: str
    cgo_type: str
    containers: dict[str, str]  # raw label -> rate suffix ("20"/"40"/"40hc"/"45")
    reefer_container_label: str | None = None  # e.g. "RD5": separate Prefix row, not a 4th slot
    reefer_prefix: str | None = None

    def suffix_for(self, container_label: str) -> str | None:
        return self.containers.get(container_label)


def load_container_map(lane_id: str) -> ContainerMap:
    path = CONTAINER_MAP_DIR / f"{lane_id.lower()}.yaml"
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    containers = {k: str(v).lower() for k, v in data["containers"].items()}
    bad = set(containers.values()) - VALID_SUFFIXES
    if bad:
        raise ValueError(f"{path}: unknown rate-column suffix(es) {bad}")
    return ContainerMap(
        prefix=data["prefix"],
        cgo_type=data["cgo_type"],
        containers=containers,
        reefer_container_label=data.get("reefer_container_label"),
        reefer_prefix=data.get("reefer_prefix"),
    )
