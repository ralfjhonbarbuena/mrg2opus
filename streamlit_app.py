"""Entry point for Streamlit Cloud (and any `streamlit run` from elsewhere).

`streamlit run` puts the SCRIPT's own folder on sys.path, not the working
directory. Pointing a deploy straight at mrg2opus/ui/app.py therefore puts
mrg2opus/ui on the path and the repo root nowhere, so the app's very first
import - `from mrg2opus.parsers import ...` - raises ModuleNotFoundError.
It only appeared to work locally because we launch from the repo root with
`python -m streamlit`, which does put the root on the path.

Keeping this file at the repo root fixes that by construction: its own
folder IS the root. It's also the filename Streamlit Cloud looks for by
default, so a deploy needs no configuration.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Belt and braces: correct already when run as the main script, and makes
# the import work anyway if some launcher sets sys.path differently.
ROOT = str(Path(__file__).resolve().parent)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from mrg2opus.ui.app import main  # noqa: E402  (must follow the sys.path fix)

main()
