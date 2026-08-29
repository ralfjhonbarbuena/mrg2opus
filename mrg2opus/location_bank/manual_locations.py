"""Location Bank entries confirmed while building the NZ1 SEA to NZBP FAK
parser, none of which any "Sample MRGs with OPUS FORMATS" PORT-PORT sheet
ever mined (bootstrap_from_samples.py's only other source). Each code/name
pair below is read directly off reference/2_OPUS/41 and 42's real RATES
sheets - byte-identical across both reference weeks - the same
already-confirmed-ground-truth standard bootstrap_from_samples.py itself
uses, just sourced from reference/ instead of the Sample MRGs folder.

Run as: python -m mrg2opus.location_bank.manual_locations
"""
from __future__ import annotations

from mrg2opus.location_bank.known_aliases import KNOWN_ALIASES
from mrg2opus.location_bank.models import LocationRecord
from mrg2opus.location_bank.store import LocationBankStore

# code -> (primary_name, country)
MANUAL_LOCATIONS: dict[str, tuple[str, str | None]] = {
    "MYKUA": ("KUANTAN", "MY"),
    "BNMUA": ("MUARA", "BN"),
    "IDMAK": ("MAKASSAR", "ID"),
    "IDPDG": ("PADANG", "ID"),
    "IDBDJ": ("BANJARMASIN", "ID"),
    "IDBPN": ("BALIKPAPAN", "ID"),
    "IDPNK": ("PONTIANAK", "ID"),
    "IDBTM": ("BATAM", "ID"),
    "IDSRI": ("SAMARINDA", "ID"),
    "PHSFS": ("SUBIC BAY", "PH"),
    "PHCGY": ("CAGAYAN DE ORO", "PH"),
    "PKBQM": ("MUHAMMAD BIN QASIM", "PK"),
}


def seed(store: LocationBankStore | None = None) -> int:
    store = store or LocationBankStore()
    for code, (name, country) in MANUAL_LOCATIONS.items():
        store.upsert_location(
            LocationRecord(code=code, primary_name=name, country=country, source="manual_override")
        )
    # bootstrap_from_samples.bootstrap() only wires KNOWN_ALIASES entries in
    # when it re-mines the Sample MRGs folder from scratch; re-add here too
    # so a targeted re-run of this module alone stays sufficient.
    for alias, code in KNOWN_ALIASES.items():
        store.add_alias(alias, code, source="manual_override")
    return len(MANUAL_LOCATIONS)


if __name__ == "__main__":
    n = seed()
    print(f"Seeded {n} manual location record(s)")
