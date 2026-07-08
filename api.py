"""
FastAPI wrapper around oracle_patch_notes_to_excel.py.

Exposes the scraper as a REST endpoint so external callers — e.g. an
Oracle Fusion AI Agent Studio tool/function — can trigger a scrape by
passing parameters and get back a link to the generated Excel workbook.
"""

import os
import re
import secrets
import uuid
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from openpyxl import Workbook

from oracle_patch_notes_to_excel import (
    BASE,
    MODULE_NAME_MAP,
    build_summary_sheet,
    discover_modules,
    process_module,
)

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "generated"))
OUTPUT_DIR.mkdir(exist_ok=True)

FILE_NAME_RE = re.compile(r"[\w\-.]+\.xlsx")

API_USERNAME = os.environ.get("API_USERNAME", "admin")
API_PASSWORD = os.environ.get("API_PASSWORD", "admin")

security = HTTPBasic()


def require_auth(credentials: HTTPBasicCredentials = Depends(security)):
    valid_user = secrets.compare_digest(credentials.username, API_USERNAME)
    valid_pass = secrets.compare_digest(credentials.password, API_PASSWORD)
    if not (valid_user and valid_pass):
        raise HTTPException(
            status_code=401,
            detail="Invalid credentials",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


app = FastAPI(
    title="Oracle Patch Notes to Excel API",
    description=(
        "Fetches Oracle Fusion Cloud readiness / \"What's New\" content from "
        "Oracle Help Center and returns a link to a generated Excel workbook."
    ),
    version="1.0.0",
)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/")
def root(_: str = Depends(require_auth)):
    """Auth-gated root route — connection-test pings from external tools land here."""
    return {"service": "oracle-patch-notes-api", "status": "ok"}


@app.get("/modules")
def list_modules(_: str = Depends(require_auth)):
    """Known module names usable with the `module` parameter, e.g. 'inventory', 'payroll'."""
    return {"modules": sorted(MODULE_NAME_MAP.keys())}


@app.get("/generate")
def generate(
    request: Request,
    _: str = Depends(require_auth),
    release: str = Query(..., description="Release code, e.g. 26b, 26a"),
    module: Optional[str] = Query(None, description="Known module name — see GET /modules"),
    url: Optional[str] = Query(None, description="Explicit module index URL path, e.g. scm/26b/mfg26b/index.html"),
    all_modules: bool = Query(False, alias="all", description="Fetch all modules across all pillars for the release"),
    max_pages: int = Query(60, ge=1, le=200, description="Max pages to follow per module"),
):
    """
    Scrape one module, an explicit URL, or every module for a release, and
    write the results to an Excel workbook. Returns JSON with a download_url
    for the generated file (see GET /download/{file_name}).
    """
    selected = [x for x in (module, url, all_modules) if x]
    if len(selected) != 1:
        raise HTTPException(400, "Provide exactly one of: module, url, all=true")

    release = release.strip().lower()
    wb = Workbook()
    wb.remove(wb.active)
    summary_rows = []

    try:
        if all_modules:
            modules = discover_modules(release)
            if not modules:
                raise HTTPException(404, f"No modules discovered for release '{release}'")
            for index_url, slug in modules:
                summary_rows.append(process_module(wb, index_url, slug, release, max_pages))
        elif module:
            name = module.strip().lower()
            path = MODULE_NAME_MAP.get(name)
            if not path:
                close = [k for k in MODULE_NAME_MAP if name in k or k in name]
                hint = f" Did you mean: {', '.join(close)}?" if close else " See GET /modules for valid names."
                raise HTTPException(400, f"Unknown module '{module}'.{hint}")
            path = re.sub(r"26b", release, path, flags=re.I)
            index_url = f"{BASE}/{path}"
            summary_rows.append(process_module(wb, index_url, name, release, max_pages))
        else:  # url
            index_url = url if url.startswith("http") else f"{BASE}/{url}"
            summary_rows.append(process_module(wb, index_url, url.rstrip("/").split("/")[-2], release, max_pages))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"Failed to fetch/parse Oracle documentation: {e}")

    build_summary_sheet(wb, release, summary_rows)

    file_name = f"oracle_fusion_{release}_readiness_{uuid.uuid4().hex[:8]}.xlsx"
    wb.save(OUTPUT_DIR / file_name)

    total_features = sum(r[1] for r in summary_rows if isinstance(r[1], int))
    base_url = str(request.base_url).rstrip("/")

    return {
        "status": "success",
        "release": release.upper(),
        "modules_processed": len(summary_rows),
        "total_features": total_features,
        "file_name": file_name,
        "download_url": f"{base_url}/download/{file_name}",
    }


@app.get("/download/{file_name}")
def download(file_name: str, _: str = Depends(require_auth)):
    if not FILE_NAME_RE.fullmatch(file_name):
        raise HTTPException(400, "Invalid file name")
    path = OUTPUT_DIR / file_name
    if not path.is_file():
        raise HTTPException(404, "File not found or already cleaned up")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_name,
    )
