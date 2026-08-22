"""SQLite-backed Location Bank: indexed exact lookup by code, plus a full
listing for the fuzzy matcher to score against. Chosen over a flat JSON/
parquet file because Step 3's manual-override review needs a safe write
path, and SQLite is a trivial future migration target to Postgres if the
app moves to a shared deployment.
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path

from mrg2opus.location_bank.models import LocationRecord

DEFAULT_DB_PATH = Path(__file__).resolve().parents[2] / "data" / "location_bank.sqlite3"

SCHEMA = """
CREATE TABLE IF NOT EXISTS locations (
    code TEXT PRIMARY KEY,
    primary_name TEXT NOT NULL,
    country TEXT,
    subdivision TEXT,
    is_unlocode INTEGER NOT NULL DEFAULT 0,
    source TEXT NOT NULL,
    confidence REAL NOT NULL DEFAULT 1.0
);
CREATE TABLE IF NOT EXISTS location_aliases (
    alias TEXT NOT NULL,
    code TEXT NOT NULL REFERENCES locations(code),
    source TEXT NOT NULL,
    PRIMARY KEY (alias, code)
);
"""


class LocationBankStore:
    def __init__(self, db_path: Path | str = DEFAULT_DB_PATH):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            conn.executescript(SCHEMA)

    @contextmanager
    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def upsert_location(self, record: LocationRecord, allow_downgrade: bool = False) -> None:
        """Insert or update a location. Unless allow_downgrade, a higher-authority
        source (unlocode > manual_override > sample_mined) is not overwritten by
        a lower one for primary_name/country/subdivision.
        """
        rank = {"unlocode": 3, "manual_override": 2, "sample_mined": 1}
        with self._connect() as conn:
            existing = conn.execute(
                "SELECT source FROM locations WHERE code = ?", (record.code,)
            ).fetchone()
            if existing and not allow_downgrade:
                if rank.get(existing[0], 0) > rank.get(record.source, 0):
                    return
            conn.execute(
                """
                INSERT INTO locations (code, primary_name, country, subdivision, is_unlocode, source, confidence)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    primary_name=excluded.primary_name,
                    country=excluded.country,
                    subdivision=excluded.subdivision,
                    is_unlocode=excluded.is_unlocode,
                    source=excluded.source,
                    confidence=excluded.confidence
                """,
                (
                    record.code,
                    record.primary_name,
                    record.country,
                    record.subdivision,
                    int(record.is_unlocode),
                    record.source,
                    record.confidence,
                ),
            )

    def add_alias(self, alias: str, code: str, source: str = "sample_mined") -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO location_aliases (alias, code, source) VALUES (?, ?, ?)",
                (alias.strip(), code, source),
            )

    def get_by_code(self, code: str) -> LocationRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT code, primary_name, country, subdivision, is_unlocode, source, confidence "
                "FROM locations WHERE code = ?",
                (code,),
            ).fetchone()
        if row is None:
            return None
        return LocationRecord(
            code=row[0], primary_name=row[1], country=row[2], subdivision=row[3],
            is_unlocode=bool(row[4]), source=row[5], confidence=row[6],
        )

    def get_by_alias(self, alias: str) -> list[LocationRecord]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT l.code, l.primary_name, l.country, l.subdivision, l.is_unlocode, l.source, l.confidence "
                "FROM location_aliases a JOIN locations l ON a.code = l.code "
                "WHERE a.alias = ?",
                (alias.strip(),),
            ).fetchall()
        return [
            LocationRecord(
                code=r[0], primary_name=r[1], country=r[2], subdivision=r[3],
                is_unlocode=bool(r[4]), source=r[5], confidence=r[6],
            )
            for r in rows
        ]

    def all_locations(self) -> list[LocationRecord]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT code, primary_name, country, subdivision, is_unlocode, source, confidence FROM locations"
            ).fetchall()
        return [
            LocationRecord(
                code=r[0], primary_name=r[1], country=r[2], subdivision=r[3],
                is_unlocode=bool(r[4]), source=r[5], confidence=r[6],
            )
            for r in rows
        ]

    def count(self) -> int:
        with self._connect() as conn:
            return conn.execute("SELECT COUNT(*) FROM locations").fetchone()[0]
