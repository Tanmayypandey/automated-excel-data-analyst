"""FastAPI Backend Server for Automated Excel Data Analyst.

Exposes REST endpoints for file upload, analysis pipeline execution,
downloading generated Excel workbooks, and serving the interactive web dashboard.
"""

import os
import re
import uuid
import shutil
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware

from src.pipeline import run_pipeline


BASE_DIR = Path(__file__).resolve().parent

DATA_INPUT_DIR = BASE_DIR / "data" / "input"
DATA_OUTPUT_DIR = BASE_DIR / "data" / "output"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

# Ensure directories exist
os.makedirs(DATA_INPUT_DIR, exist_ok=True)
os.makedirs(DATA_OUTPUT_DIR, exist_ok=True)

app = FastAPI(
    title="Automated Excel Data Analyst",
    description="Local Autonomous Excel & CSV Analytics and Dashboard Builder",
    version="1.0.0"
)

# ---------------------------------------------------------
# CORS CONFIGURATION
# ---------------------------------------------------------
# For now "*" allows your future Vercel frontend to call Render.
# Later, after Vercel gives you the final URL, we can restrict this
# to only that exact Vercel domain.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------
# STATIC FILES AND TEMPLATES
# ---------------------------------------------------------
app.mount(
    "/static",
    StaticFiles(directory=str(STATIC_DIR)),
    name="static"
)

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


@app.get("/", response_class=HTMLResponse)
async def index_view(request: Request):
    accept = request.headers.get("accept", "")

    if "text/html" in accept or "*/*" in accept or not accept:
        return FileResponse(
            str(TEMPLATES_DIR / "index.html"),
            media_type="text/html"
        )

    return JSONResponse({
        "status": "running",
        "message": "Automated Excel Data Analyst API",
        "supported_files": [".xlsx", ".csv"]
    })


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "service": "Automated Excel Data Analyst"
    }


@app.post("/analyze")
async def analyze_file(file: UploadFile = File(...)):
    """Upload dataset (.xlsx or .csv), run analysis pipeline,
    and produce the .xlsx dashboard.
    """

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file uploaded."
        )

    orig_filename = file.filename

    _, ext = os.path.splitext(orig_filename)
    ext_lower = ext.lower()

    if ext_lower not in [".xlsx", ".csv"]:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported file extension '{ext}'. "
                "Only .xlsx and .csv files are supported."
            )
        )

    # Generate safe unique filenames
    clean_stem = re.sub(
        r"[^\w\-.]",
        "_",
        Path(orig_filename).stem
    )

    file_id = uuid.uuid4().hex[:10]

    saved_input_name = (
        f"{clean_stem}_{file_id}{ext_lower}"
    )

    output_xlsx_name = (
        f"{clean_stem}_analyzed_{file_id}.xlsx"
    )

    input_path = str(
        DATA_INPUT_DIR / saved_input_name
    )

    output_path = str(
        DATA_OUTPUT_DIR / output_xlsx_name
    )

    try:
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(
                file.file,
                buffer
            )

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to save uploaded file: {str(e)}"
        )

    finally:
        file.file.close()

    # Run analytics pipeline
    try:
        result = run_pipeline(
            input_file_path=input_path,
            output_file_path=output_path
        )

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Pipeline execution error: {str(e)}"
        )

    if result.get("status") != "success":
        error_msg = (
            result.get("error")
            or "Unknown error occurred during analysis."
        )

        raise HTTPException(
            status_code=422,
            detail=error_msg
        )

    return {
        "status": "success",
        "message": "Data analysis completed successfully.",
        "original_filename": orig_filename,
        "input_type": ext_lower,
        "selected_sheet": result.get(
            "selected_sheet"
        ),
        "charts_generated": len(
            result.get(
                "selected_charts",
                []
            )
        ),
        "output_filename": output_xlsx_name,
        "download_url": (
            f"/download/{output_xlsx_name}"
        ),
        "kpis": result.get(
            "selected_kpis",
            []
        ),
        "selected_charts": result.get(
            "selected_charts",
            []
        ),
        "cleaning_report": result.get(
            "cleaning_report",
            {}
        ),
        "validation_report": result.get(
            "validation_report",
            {}
        )
    }


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download generated Excel report with path traversal protection."""

    safe_filename = os.path.basename(
        filename
    )

    target_path = (
        DATA_OUTPUT_DIR /
        safe_filename
    )

    if (
        not target_path.exists()
        or not target_path.is_file()
    ):
        raise HTTPException(
            status_code=404,
            detail="Requested report file not found."
        )

    return FileResponse(
        path=str(target_path),
        filename=safe_filename,
        media_type=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )