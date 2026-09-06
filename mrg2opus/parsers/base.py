from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any, ClassVar

from openpyxl.workbook import Workbook

from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import OpusRowSet


@dataclass
class RawExtraction:
    """Intermediate representation between 'raw Excel' and 'OPUS rows'.

    Deliberately loose (dict-of-tables) rather than a fixed schema, since
    each lane's raw layout differs; to_opus_rows() is where lane-specific
    knowledge turns this into the strict OpusRowSet contract.
    """

    tables: dict[str, Any] = field(default_factory=dict)


class UnclassifiedMRGError(Exception):
    pass


class BaseMRGParser(ABC):
    lane_id: ClassVar[str]

    # Per-lane real-filing sheet-name overrides, keyed by OpusRowSet field
    # name (e.g. {"route_notes": "ROUTE NOTE"}) - most lanes' route notes
    # file as "RN" (confirmed against LAWC's real ground truth), but a lane
    # can override when its own ground truth uses something else (TAD
    # FILING's real sheets are literally named "ROUTE NOTE" - see
    # excel_io/writer.py's _sheet_names_for_suffix for the base names this
    # overrides). Empty by default; step4_export.py reads this off the
    # active parser class and passes it to write_opus_workbook_multi.
    SHEET_NAME_OVERRIDES: ClassVar[dict[str, str]] = {}

    # What the REAL filings call each scope's sheets, where that differs
    # from the name we write. Read-side only: Compare looks a sheet up in
    # a reference workbook by this name, and nothing else consults it.
    #
    # It used to name our own output too, which is how one lane's export
    # ended up with its four scopes' surcharge sheets called "SRCHG",
    # "CMDT NOTE-AMW", "AEW SRCHG" and "AMW SRCHG" (user-reported,
    # 2026-09-06). Those are three separate real workbooks' conventions,
    # and they only collide because we write one workbook where the team
    # files three - so our side now uses the same "{base}-{scope}" scheme
    # for every scope of every lane, and the filings' own names live here
    # for finding them again.
    #
    # {scope: {OpusRowSet field name: the sheet name in that scope's real
    # filing}}. A field not listed is looked up under the normal name.
    REFERENCE_SHEET_NAMES: ClassVar[dict[str, dict[str, str]]] = {}

    @abstractmethod
    def parse_raw(self, wb: Workbook) -> RawExtraction:
        """Excel -> intermediate tables, with exclusion (strikethrough/fill)
        already applied and merged headers already flattened."""

    @abstractmethod
    def to_opus_rows(self, raw: RawExtraction, config: MappingProfile) -> OpusRowSet:
        """Intermediate tables -> populated OPUS row models."""

    def run(self, wb: Workbook, config: MappingProfile | None = None) -> OpusRowSet:
        config = config or MappingProfile()
        raw = self.parse_raw(wb)
        return self.to_opus_rows(raw, config)

    def run_multi(self, wb: Workbook, config: MappingProfile | None = None) -> dict[str, OpusRowSet]:
        """For lanes with sub-lanes sharing one workbook (e.g. EAF's
        TZDAR/KEMBA, each with their own suffixed OPUS sheets), override
        this to return {suffix: OpusRowSet}. Default wraps run() as the
        single unsuffixed output, so single-output lanes need no changes."""
        return {"": self.run(wb, config)}
