"""
Google Sheets sync logic.

Syncs parsed rows to a Google Sheet using a service account:
  - Updates rows that already exist (matched by promo number in col A)
  - Appends rows that are new
  - Removes rows in the sheet not present in the parsed doc

Column layout matches the xlsx writer (A–G).
"""

import json
import os

import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

HEADERS = ["", "Name", "Promo Number", "Promo Name", "Cue", "Notes", "Character"]


# ── Config / auth ─────────────────────────────────────────────────────────────

def load_config() -> dict:
    with open(CONFIG_PATH) as f:
        return json.load(f)


def get_worksheet(sheet_type: str) -> gspread.Worksheet:
    config = load_config()
    sheet_cfg = config["sheets"].get(sheet_type)
    if not sheet_cfg:
        raise ValueError(f"Unknown show type: '{sheet_type}'. Must be 'PPV' or 'FN'.")

    creds_path = sheet_cfg["credentials_file"]
    if not os.path.isfile(creds_path):
        raise FileNotFoundError(
            f"Credentials file not found: {creds_path}\n"
            "Download a service account JSON key and place it at that path."
        )

    spreadsheet_id = sheet_cfg["spreadsheet_id"]
    if spreadsheet_id.startswith("YOUR_"):
        raise ValueError(
            f"Spreadsheet ID not configured for '{sheet_type}'. "
            "Edit config.json and set the correct spreadsheet_id."
        )

    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(spreadsheet_id).worksheet("Sheet1")


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_num(val) -> str:
    """
    Normalise promo numbers for comparison.
    Handles floats from Excel ('1.0' -> '1') and date-misinterpreted variants.
    """
    s = str(val).strip()
    # Float-looking: '1.0' -> '1'
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except (ValueError, OverflowError):
        pass
    return s


def row_to_values(row: dict, row_idx: int) -> list:
    """Return a 7-element list for columns A–G."""
    num = row["number"]
    name = row["name"]
    cue = row["cue"]
    return [num, name, num, name, cue, "", f"=LEN(E{row_idx})"]


# ── Sync ──────────────────────────────────────────────────────────────────────

def sync_rows(rows: list[dict], sheet_type: str) -> dict:
    """
    Sync parsed rows to the named Google Sheet.

    Returns:
        {
          "added":   [list of names],
          "updated": [list of names],
          "removed": [list of names],
        }
    """
    ws = get_worksheet(sheet_type)

    # Read the full sheet (row 1 = header, rows 2+ = data)
    all_values = ws.get_all_values()

    # Ensure header row exists
    if not all_values:
        ws.append_row(HEADERS)
        all_values = [HEADERS]

    data_rows = all_values[1:]  # 0-indexed list of row content

    # Map: normalized_number -> sheet_row_1based
    existing: dict[str, int] = {}
    for i, row_vals in enumerate(data_rows, start=2):
        num = normalize_num(row_vals[0]) if row_vals and row_vals[0] else ""
        if num:
            existing[num] = i

    # Map: normalized_number -> parsed_row (preserving doc order)
    target: dict[str, dict] = {normalize_num(r["number"]): r for r in rows}

    added: list[str] = []
    updated: list[str] = []
    removed: list[str] = []

    # 1. Batch-update rows that already exist in the sheet
    batch_data = []
    for num, row in target.items():
        if num in existing:
            sheet_row = existing[num]
            batch_data.append({
                "range": f"A{sheet_row}:G{sheet_row}",
                "values": [row_to_values(row, sheet_row)],
            })
            updated.append(row["name"])

    if batch_data:
        ws.batch_update(batch_data, value_input_option="USER_ENTERED")

    # 2. Delete sheet rows whose promo numbers are not in the doc
    #    Process bottom-to-top so row indices don't shift during deletion.
    to_delete = sorted(
        [idx for num, idx in existing.items() if num not in target],
        reverse=True,
    )
    for sheet_row in to_delete:
        # Grab name from cached data before deleting
        cached = data_rows[sheet_row - 2]
        removed.append(cached[1] if len(cached) > 1 else "?")
        ws.delete_rows(sheet_row)

    # 3. Append rows that are new (not in the sheet at all)
    rows_to_append = [r for num, r in target.items() if num not in existing]
    if rows_to_append:
        # Re-count rows after deletions to get correct row numbers for LEN formula
        current_count = len(ws.get_all_values())
        append_values = []
        for i, row in enumerate(rows_to_append, start=current_count + 1):
            append_values.append(row_to_values(row, i))
            added.append(row["name"])
        ws.append_rows(append_values, value_input_option="USER_ENTERED")

    # 4. Standardise formatting across all data rows
    total_rows = len(ws.get_all_values())
    if total_rows >= 2:
        # Arial 10pt everywhere, no bold by default
        ws.format(f"A2:G{total_rows}", {
            "textFormat": {"fontFamily": "Arial", "fontSize": 10, "bold": False},
        })
        # Columns A, C: Georgia font, center both axes
        for col in ("A", "C"):
            ws.format(f"{col}2:{col}{total_rows}", {
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Georgia", "fontSize": 10, "bold": False},
            })
        # Column F: center both axes
        ws.format(f"F2:F{total_rows}", {
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        })
        # Columns B, D, G: bold + center both axes
        for col in ("B", "D", "G"):
            ws.format(f"{col}2:{col}{total_rows}", {
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Arial", "fontSize": 10, "bold": True},
            })
        # Column E (Cue): wrap text, top + left
        ws.format(f"E2:E{total_rows}", {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "TOP",
            "wrapStrategy": "WRAP",
        })
        # Auto-resize all data rows to fit wrapped content in column E
        # Also auto-resize columns B (idx 1) and D (idx 3) to fit full text
        ws.spreadsheet.batch_update({"requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": 1,
                        "endIndex": total_rows,
                    }
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,  # Column B
                        "endIndex": 2,
                    }
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": 3,  # Column D
                        "endIndex": 4,
                    }
                }
            },
        ]})

    return {"added": added, "updated": updated, "removed": removed}


# ── Config validation (called by the app on startup) ─────────────────────────

def validate_config() -> list[str]:
    """Return a list of warning strings for any misconfigured show types."""
    warnings = []
    try:
        config = load_config()
    except Exception as e:
        return [f"config.json unreadable: {e}"]

    for show_type, sheet_cfg in config.get("sheets", {}).items():
        sid = sheet_cfg.get("spreadsheet_id", "")
        creds = sheet_cfg.get("credentials_file", "")
        if sid.startswith("YOUR_"):
            warnings.append(f"{show_type}: spreadsheet_id not set in config.json")
        if not os.path.isfile(creds):
            warnings.append(f"{show_type}: credentials file not found ({creds})")

    return warnings
