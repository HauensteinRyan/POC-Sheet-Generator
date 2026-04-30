# POC Sheet Generator

Converts UFC promo script Word docs (.docx) into Excel spreadsheets (.xlsx) matching the PPVPOC format.

## Cloudflare deployment

This repo now includes a Cloudflare Worker migration that keeps the same browser-facing routes as the FastAPI app:

- `GET /`, `GET /login`, `GET /view`
- `POST /parse`
- `POST /upload`
- `POST /download-rows`
- `POST /sync`
- `GET /config-status`

Cloudflare serves the files in `static/` through the Worker assets binding, while `src/worker.ts` handles auth, `.docx` parsing, `.xlsx` generation, and Google Sheets sync.

### Local Worker dev

```bash
npm install
cp .dev.vars.example .dev.vars
# edit .dev.vars with real secret values
npm run dev:worker
```

### Cloudflare secrets

Set these in Cloudflare before production deploys:

```bash
npx wrangler secret put SESSION_SECRET
npx wrangler secret put APP_USERS_JSON
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_JSON
```

`APP_USERS_JSON` should look like:

```json
{"admin":"change-this-password"}
```

`GOOGLE_SERVICE_ACCOUNT_JSON` should be the full service account JSON key as a single secret value. The target spreadsheet ID is configured inside the app by pasting a Google Sheets URL or spreadsheet ID into the Google Sheet field; the browser remembers separate PPV and FN targets.

### Git deploy setup

In Cloudflare, create a Worker connected to the GitHub repo:

1. Workers & Pages -> Create -> Import a repository.
2. Select `prusik-haulbag/poc-sheet-generator`.
3. Use `npm install` as the install command.
4. Use `npm run deploy` as the deploy command if Cloudflare asks for one.
5. Set the three secrets above in the Worker settings.

After that, pushes to the connected branch can deploy the Worker.

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
