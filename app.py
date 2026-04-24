"""
FastAPI server — localhost only for now.

GET  /        → frontend HTML
POST /parse   → returns parsed rows as JSON (for doc preview)
POST /upload  → returns generated .xlsx as a file download
POST /sync    → syncs parsed rows to the configured Google Sheet
GET  /config-status → reports whether PPV/FN credentials are configured
"""

import io
import os
import tempfile

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from parser import parse_doc
from writer import write_xlsx
from sheets import sync_rows, validate_config

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")


def _save_upload(file: UploadFile, content: bytes) -> str:
    """Write upload bytes to a temp .docx file and return the path."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(content)
        return tmp.name


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
    file: UploadFile = File(...),
    show_type: str = Form(...),
):
    """
    Parse .docx then sync rows to the Google Sheet for the given show_type
    ('PPV' or 'FN').  Returns a summary of added / updated / removed rows.
    """
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    show_type = show_type.upper()
    if show_type not in ("PPV", "FN"):
        raise HTTPException(status_code=400, detail="show_type must be 'PPV' or 'FN'.")

    content = await file.read()
    tmp_path = _save_upload(file, content)
    try:
        rows = parse_doc(tmp_path)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to parse document: {e}")
    finally:
        os.unlink(tmp_path)

    try:
        result = sync_rows(rows, show_type)
    except (FileNotFoundError, ValueError) as e:
        raise HTTPException(status_code=503, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Google Sheets error: {e}")

    return {
        "show_type": show_type,
        "added":   result["added"],
        "updated": result["updated"],
        "removed": result["removed"],
        "total_rows": len(rows),
    }
