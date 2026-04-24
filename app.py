"""
FastAPI server — localhost only for now.

GET  /        → frontend HTML
POST /parse   → returns parsed rows as JSON (for doc preview)
POST /upload  → returns generated .xlsx as a file download
POST /sync    → syncs parsed rows to the configured Google Sheet
GET  /config-status → reports whether PPV/FN credentials are configured
GET  /login   → login page
POST /auth/login → validate credentials, set session cookie
GET  /auth/logout → clear session cookie
"""

import io
import os
import tempfile

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from auth import COOKIE_NAME, make_token, verify_credentials, verify_token
from parser import parse_doc
from sheets import sync_rows, validate_config
from writer import write_xlsx

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

# Paths that don't require a session
_EXEMPT = {"/login", "/auth/login", "/auth/logout"}


@app.middleware("http")
async def require_auth(request: Request, call_next):
    if request.url.path in _EXEMPT or request.url.path.startswith("/static/"):
        return await call_next(request)
    token = request.cookies.get(COOKIE_NAME)
    if not verify_token(token):
        return RedirectResponse("/login", status_code=303)
    return await call_next(request)


# ── Auth routes ───────────────────────────────────────────────────────────────

@app.get("/login")
async def login_page():
    return FileResponse("static/login.html")


@app.post("/auth/login")
async def do_login(username: str = Form(...), password: str = Form(...)):
    if not verify_credentials(username, password):
        return RedirectResponse("/login?error=1", status_code=303)
    token = make_token(username)
    response = RedirectResponse("/", status_code=303)
    response.set_cookie(COOKIE_NAME, token, httponly=True, samesite="lax")
    return response


@app.get("/auth/logout")
async def logout():
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(COOKIE_NAME)
    return response


# ── Helpers ───────────────────────────────────────────────────────────────────

def _save_upload(file: UploadFile, content: bytes) -> str:
    """Write upload bytes to a temp .docx file and return the path."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(content)
        return tmp.name


# ── App routes ────────────────────────────────────────────────────────────────

@app.get("/")
async def index():
    return FileResponse("static/index.html")


@app.get("/config-status")
async def config_status():
    """Let the frontend know which show types are ready to sync."""
    warnings = validate_config()
    return {"warnings": warnings}


@app.post("/parse")
async def parse(file: UploadFile = File(...)):
    """Return parsed rows as JSON for the doc preview panel."""
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    content = await file.read()
    tmp_path = _save_upload(file, content)
    try:
        rows = parse_doc(tmp_path)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to parse document: {e}")
    finally:
        os.unlink(tmp_path)

    return {"rows": rows, "count": len(rows)}


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    """Parse .docx and return a .xlsx file download."""
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    content = await file.read()
    tmp_path = _save_upload(file, content)
    try:
        rows = parse_doc(tmp_path)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to parse document: {e}")
    finally:
        os.unlink(tmp_path)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
        tmp_out_path = tmp_out.name
    try:
        write_xlsx(rows, tmp_out_path)
        with open(tmp_out_path, "rb") as f:
            xlsx_bytes = f.read()
    finally:
        os.unlink(tmp_out_path)

    base_name = os.path.splitext(file.filename)[0]
    out_filename = f"{base_name}_output.xlsx"

    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_filename}"'},
    )


@app.post("/sync")
async def sync(
    rows: str = Form(...),
    show_type: str = Form(...),
):
    import json as _json

    show_type = show_type.upper()
    if show_type not in ("PPV", "FN"):
        raise HTTPException(status_code=400, detail="show_type must be 'PPV' or 'FN'.")

    try:
        row_list = _json.loads(rows)
    except Exception:
        raise HTTPException(status_code=422, detail="Invalid rows JSON.")

    try:
        result = sync_rows(row_list, show_type)
    except (FileNotFoundError, ValueError) as e:
        raise HTTPException(status_code=503, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Google Sheets error: {e}")

    return {
        "show_type": show_type,
        "added":   result["added"],
        "updated": result["updated"],
        "removed": result["removed"],
        "total_rows": len(row_list),
    }


class DownloadRequest(BaseModel):
    rows: list[dict]
    filename: str = "output"


@app.post("/download-rows")
async def download_rows(payload: DownloadRequest):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        write_xlsx(payload.rows, tmp_path)
        with open(tmp_path, "rb") as f:
            xlsx_bytes = f.read()
    finally:
        os.unlink(tmp_path)

    out_filename = f"{payload.filename}_output.xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_filename}"'},
    )
