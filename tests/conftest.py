from __future__ import annotations

from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLES_DIR = REPO_ROOT / "Sample MRGs with OPUS FORMATS"
RAW_SAMPLES_DIR = REPO_ROOT / "MRGs RAW SAMPLES"


@pytest.fixture(scope="session")
def samples_dir() -> Path:
    return SAMPLES_DIR


@pytest.fixture(scope="session")
def raw_samples_dir() -> Path:
    return RAW_SAMPLES_DIR
