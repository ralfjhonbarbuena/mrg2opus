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

    # Per-SCOPE (run_multi() sub-lane key) full sheet-name overrides, for
    # the rarer case where SHEET_NAME_OVERRIDES' uniform "{base}-{suffix}"
    # tagging can't express the real naming - confirmed e.g. TAD FILING
    # AEW/AMW: AEW's own CMDT NOTE block is literally named "SRCHG" while
    # AMW's uses the standard "CMDT NOTE" name, and AEW's ARBS sheet is
    # "AEW ARBS" (name PREFIXED, no hyphen) while AMW's is bare "ORIGIN
    # ARBS" (no scope tag at all) - no single suffix-tag rule covers both.
    # {scope: {OpusRowSet field name: full sheet name}} - a scope present
    # here uses these names verbatim instead of the tagged default/
    # SHEET_NAME_OVERRIDES for the fields it lists; fields it doesn't list
    # still fall back to the normal tagged naming.
    SCOPED_SHEET_NAME_OVERRIDES: ClassVar[dict[str, dict[str, str]]] = {}

    @classmethod
    @abstractmethod
    def detect(cls, wb: Workbook) -> float:
        """Return a confidence score in [0, 1] that this parser matches wb."""

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
