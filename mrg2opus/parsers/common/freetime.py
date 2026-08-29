"""Static per-lane FREETIME (free-time-allowance/RFA reference table)
builders.

Confirmed against real ground truth that this sheet is NOT derived from
the raw MRG's rate data at all:
- LAWC: byte-identical (incl. EFF DT/EXP DT) across every one of 3 real
  ground-truth files seen (FAK week 1, FAK week 2, TIER 1) - a pure
  filing-wide constant.
- LAEC: identical content across every real ground-truth file seen EXCEPT
  EFF DT/EXP DT, which tracks that filing's own rate validity window (same
  pattern CMDT NOTE's own APP row uses) - and TIER 1 adds exactly one row
  (Argentina/Zarate) on top of FAK's own table. LAEC LUX has its own,
  much smaller table (BR/AR/UY/PY only, no RGN/LOC breakdown).

These constants should be updated only when a new confirmed ground-truth
sample shows the underlying reference data has actually changed (e.g. the
LAWC RFA's own 2026-09-30 expiry) - not guessed at.
"""
from __future__ import annotations

from datetime import date

from mrg2opus.schema.opus_rows import FreetimeRow

# --- LAWC ---------------------------------------------------------------

_LAWC_SHARED = dict(
    rfa_no="SINN02654A", status="Approved", tariff="CTIC",
    eff_dt=date(2023, 3, 16), exp_dt=date(2026, 9, 30),
    free_time_tier="Single", ftime_excl_sat="N", ftime_excl_sun="N", ftime_excl_hday="N",
    origin_or_dest_ct="A", remark="LWE",
    dar_no="SINBB23030029B", ver="015", approval_no="SINHQ23034238B", proposal_no="RSIN230036",
    customer_code="SG100628", customer_name="C.H. ROBINSON FREIGHT SERVICES (SINGAPORE) PTE, LTD.",
)

# (cntr_cargo, coverage_cn, free_time_total), in ground-truth row order.
_LAWC_ROWS: list[tuple[str, str, int]] = [
    ("Dry - General", "CL", 21), ("Dry - General", "PE", 21), ("Dry - General", "MX", 21),
    ("Dry - General", "CO", 20), ("Dry - General", "EC", 18), ("Dry - General", "GT", 17),
    ("Dry - General", "NI", 16), ("Dry - General", "CR", 16), ("Dry - General", "SV", 16),
    ("Dry - General", "HN", 16), ("Dry - General", "PA", 12),
    ("Reefer - General", "CL", 14), ("Reefer - General", "PE", 14), ("Reefer - General", "MX", 14),
    ("Reefer - General", "CO", 14), ("Reefer - General", "EC", 14), ("Reefer - General", "GT", 14),
    ("Reefer - General", "NI", 14), ("Reefer - General", "CR", 14), ("Reefer - General", "SV", 14),
    ("Reefer - General", "HN", 14), ("Reefer - General", "PA", 14),
]


def build_lawc_freetime() -> list[FreetimeRow]:
    """Pure constant - identical for LAWC FAK and LAWC TIER 1."""
    return [
        FreetimeRow(seq=i, cntr_cargo=cargo, coverage_cn=cn, free_time_total=total, **_LAWC_SHARED)
        for i, (cargo, cn, total) in enumerate(_LAWC_ROWS, start=1)
    ]


# --- LAEC -----------------------------------------------------------------

# (cntr_cargo, coverage_cn, coverage_rgn, coverage_loc, free_time_total,
#  origin_or_dest_ct, remark) - dates filled in at build time from that
# filing's own validity window.
_LAEC_BASE_ROWS: list[tuple[str, str, str | None, str | None, int, str | None, str]] = [
    ("Dry - General", "BR", "BRA", "BRSSZ", 26, "A", "LEW"),
    ("Dry - General", "AR", "ARG", "ARBUE", 26, "A", "LEW"),
    ("Dry - General", "UY", "URY", "UYMVD", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRNVT", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRITJ", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRIOA", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRPNG", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRSPB", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRRIO", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRRIG", 26, "A", "LEW"),
    ("Dry - General", "PY", None, None, 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRSUA", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRSSA", 26, "A", "LEW"),
    ("Dry - General", "BR", "BRA", "BRPEC", 26, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRSSZ", 18, "A", "LEW"),
    ("Dry - Dangerous", "AR", "ARG", "ARBUE", 18, "A", "LEW"),
    ("Dry - Dangerous", "UY", "URY", "UYMVD", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRNVT", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRITJ", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRIOA", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRPNG", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRSPB", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRRIO", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRRIG", 18, "A", "LEW"),
    ("Dry - Dangerous", "PY", None, None, 21, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRSUA", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRSSA", 18, "A", "LEW"),
    ("Dry - Dangerous", "BR", "BRA", "BRPEC", 18, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRSSZ", 16, "A", "LEW"),
    ("Reefer - General", "AR", None, None, 16, "A", "LEW"),
    ("Reefer - General", "UY", None, None, 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRNVT", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRITJ", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRIOA", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRPNG", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRSPB", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRRIO", 16, "A", "LEW"),
    ("Reefer - General", "BR", "BRA", "BRRIG", 16, "A", "LEW"),
    # "Reefer as Dry - Dangerous" - filed as its own Dry-Dangerous block,
    # not merged with the Reefer rows above (confirmed: cntr_cargo says
    # Dry - Dangerous, remark distinguishes it) - origin_or_dest_ct blank
    # here too, unlike every other row.
    ("Dry - Dangerous", "BR", "BRA", "BRSSZ", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "AR", "ARG", "ARBUE", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "UY", "URY", "UYMVD", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRNVT", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRITJ", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRIOA", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRPNG", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRSPB", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRRIO", 16, None, "LEW - Reefer as Dry - Dangerous"),
    ("Dry - Dangerous", "BR", "BRA", "BRRIG", 16, None, "LEW - Reefer as Dry - Dangerous"),
]

# TIER 1 adds exactly this one extra row on top of FAK's own table -
# confirmed against both real TIER 1 ground-truth weeks (21, 22).
_LAEC_TIER1_EXTRA_ROW = ("Dry - General", "AR", "ARG", "ARZAE", 18, "A", "LEW")

_LAEC_LUX_ROWS: list[tuple[str, str, int]] = [
    ("Dry - General", "BR", 18), ("Dry - General", "AR", 14),
    ("Dry - General", "UY", 18), ("Dry - General", "PY", 21),
    ("Dry - Dangerous", "BR", 18), ("Dry - Dangerous", "AR", 14),
    ("Dry - Dangerous", "UY", 18), ("Dry - Dangerous", "PY", 21),
]

_LAEC_SHARED = dict(tariff="CTIC", free_time_tier="Single", ftime_excl_sat="N", ftime_excl_sun="N", ftime_excl_hday="N")


def build_laec_freetime(variant: str, validity_start: date | None, validity_end: date | None) -> list[FreetimeRow]:
    """variant: "fak", "tier1", or "lux". Dates come from THIS filing's own
    rate validity window (unlike LAWC's, which never varies) - confirmed
    against the more-authoritative "TRADE_s copy" ground-truth files (a
    stale duplicate lacking "TRADE_s copy" in one FAK folder shows a
    long-lived 2025-01-01/2026-09-30 window instead - not used as the
    source of truth here). Unlike LAWC, FAK/TIER1's own Seq column is
    always blank - not populated here either (LUX's own Seq column IS
    populated 1..8 - see below). TIER 1's real ground truth also
    lists these same rows in a different block order (Dry-Dangerous
    "Reefer as..." variant first, alphabetized by LOC within each block) -
    a cosmetic filer artifact with no derivable rule, so this returns FAK's
    own row order plus the one extra TIER 1 row appended at the end rather
    than chasing an exact positional match (see test_parsers_laec.py's
    content-only comparison for TIER 1)."""
    if validity_start is None or validity_end is None:
        return []
    if variant == "lux":
        # Unlike FAK/TIER1 (Seq always blank), LUX's real ground truth
        # populates Seq 1..8 in row order - confirmed identically across
        # both real samples seen.
        return [
            FreetimeRow(
                seq=i, cntr_cargo=cargo, coverage_cn=cn, free_time_total=total,
                origin_or_dest_ct="A", eff_dt=validity_start, exp_dt=validity_end, **_LAEC_SHARED,
            )
            for i, (cargo, cn, total) in enumerate(_LAEC_LUX_ROWS, start=1)
        ]

    rows = list(_LAEC_BASE_ROWS)
    if variant == "tier1":
        rows.append(_LAEC_TIER1_EXTRA_ROW)
    return [
        FreetimeRow(
            cntr_cargo=cargo, coverage_cn=cn, coverage_rgn=rgn, coverage_loc=loc,
            free_time_total=total, origin_or_dest_ct=ct, remark=remark,
            eff_dt=validity_start, exp_dt=validity_end, **_LAEC_SHARED,
        )
        for cargo, cn, rgn, loc, total, ct, remark in rows
    ]
