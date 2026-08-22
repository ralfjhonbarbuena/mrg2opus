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
