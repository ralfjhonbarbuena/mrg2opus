from __future__ import annotations

from typing import Optional

from pydantic import BaseModel


class LocationRecord(BaseModel):
    code: str
    primary_name: str
    country: Optional[str] = None
    subdivision: Optional[str] = None
    is_unlocode: bool = False
    source: str = "sample_mined"  # "unlocode" | "sample_mined" | "manual_override"
    confidence: float = 1.0


class LocationMatch(BaseModel):
    code: str
    primary_name: str
    score: float
    needs_review: bool
