"""
Build the map for deployment
------------------------------
Finds the single input spreadsheet in data/ and runs generate_map.py
against it, producing postal_code_map_output/ (what firebase.json serves).

Run this explicitly before `firebase deploy` (it is NOT a firebase.json
predeploy hook here, unlike a typical Firebase Hosting setup: this build
calls paid Google Geocoding/Distance Matrix APIs, and a predeploy hook
would run it a second time on every `firebase deploy`, doubling API cost
per deploy). CI runs it as its own step; see
.github/workflows/firebase-hosting-merge.yml.

To update the live map: replace the spreadsheet in data/ (GitHub web UI:
data/ -> Add file -> Upload files, same name) and push/merge -- CI runs
this script and deploys automatically, no need to run anything by hand.

Requires GOOGLE_MAPS_API_KEY to be set in the environment (generate_map.py
reads it itself); this script doesn't touch it directly.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generate_map  # noqa: E402

DATA_DIR = Path(__file__).resolve().parent / "data"
SPREADSHEET_EXTENSIONS = (".xlsx", ".xls", ".csv")


def find_input_spreadsheet():
    candidates = sorted(
        p for p in DATA_DIR.glob("*")
        if p.suffix.lower() in SPREADSHEET_EXTENSIONS and not p.name.startswith("~$")
    )
    if not candidates:
        sys.exit(
            f"ERROR: no spreadsheet found in {DATA_DIR} "
            f"(expected one of {', '.join(SPREADSHEET_EXTENSIONS)})."
        )
    if len(candidates) > 1:
        sys.exit(
            f"ERROR: several spreadsheets found in {DATA_DIR}: "
            f"{', '.join(p.name for p in candidates)} -- keep only one."
        )
    return candidates[0]


if __name__ == "__main__":
    spreadsheet = find_input_spreadsheet()
    print(f"Building map from {spreadsheet}...")
    generate_map.generate_map_from_excel(str(spreadsheet))
