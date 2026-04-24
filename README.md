# POC Sheet Generator

Converts UFC promo script Word docs (.docx) into Excel spreadsheets (.xlsx) matching the PPVPOC format.

## Setup (first time only)

```bash
cd "POC Sheet Generator"
python3 -m venv venv
venv/bin/pip install -r requirements.txt
```

## Run the web app

```bash
cd "POC Sheet Generator"
venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000
```

Open **http://127.0.0.1:8000** in your browser, drop a `.docx` file, click **Convert to Excel**.

## Run from the command line

```bash
venv/bin/python main.py "path/to/script.docx"
# Output saved as path/to/script_output.xlsx

# Or specify output path:
venv/bin/python main.py "path/to/script.docx" "path/to/output.xlsx"
```

## Output columns

| Col | Content |
|-----|---------|
| A   | Promo Number (e.g. `1`, `1-1`, `10`, `10-1`) |
| B   | Name (original capitalisation from doc) |
| C   | Promo Number (duplicate of A) |
| D   | Promo Name (duplicate of B) |
| E   | Cue (uppercased) |
| F   | Notes |
| G   | `=LEN(E{row})` formula |

## Doc formatting rules

- Section headers must follow: `#N – Title` (e.g. `#1 – UFC 327 TONIGHT`)
- Variants (`ALT READ`, `Prelim read`, `Main Card read`) create sub-rows numbered `N-1`, `N-2`, etc.
- `PHONETIC – ...` lines are appended to the end of the Cue
- Sections with no spoken copy (e.g. `NO VO – N/A`) still get a row
