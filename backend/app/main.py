"""FastAPI application: upload, analyse/convert, regenerate with overrides, download."""
from __future__ import annotations

import asyncio
import logging
import zipfile
from contextlib import asynccontextmanager
from typing import Optional

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from .config import Settings
from .errors import UserError
from .jobs import Job, JobStore
from .processor import DATE_ORDERS, OUTPUT_FORMATS, OVERRIDE_TYPES, Options, process_file

log = logging.getLogger("data_format_fixer")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ALLOWED = {".xlsx": "xlsx", ".csv": "csv"}


class RegenerateRequest(BaseModel):
    date_order: str = "auto"
    output_format: str = "yyyy-mm-dd"
    overrides: dict[str, dict[str, str]] = {}


def _options(date_order: str, output_format: str) -> Options:
    if date_order not in DATE_ORDERS:
        raise UserError("Unknown date order. Use auto, dmy or mdy.", 422)
    if output_format not in OUTPUT_FORMATS:
        raise UserError("Unknown output date format.", 422)
    return Options(date_order, output_format)


def _clean_overrides(raw: dict[str, dict[str, str]]) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for sheet, cols in raw.items():
        for col, typ in cols.items():
            if typ in ("", "Auto"):
                continue
            if typ not in OVERRIDE_TYPES or not col.isdigit():
                raise UserError("Invalid column type override.", 422)
            out.setdefault(sheet, {})[col] = typ
    return out


def create_app(settings: Optional[Settings] = None) -> FastAPI:
    settings = settings or Settings.from_env()
    store = JobStore(settings.job_ttl_minutes, settings.max_active_jobs)

    async def janitor() -> None:
        while True:
            await asyncio.sleep(60)
            store.purge_expired()

    @asynccontextmanager
    async def lifespan(_: FastAPI):
        task = asyncio.create_task(janitor())
        yield
        task.cancel()
        store.close()

    app = FastAPI(title="Data Format Fixer", lifespan=lifespan, docs_url=None, redoc_url=None)
    app.state.store = store
    app.state.settings = settings
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origins,
        allow_methods=["GET", "POST", "DELETE"],
        allow_headers=["Content-Type"],
        expose_headers=["Content-Disposition"],
    )

    @app.exception_handler(UserError)
    async def user_error(_: Request, exc: UserError):
        return JSONResponse({"detail": exc.message}, status_code=exc.status_code)

    @app.exception_handler(RequestValidationError)
    async def validation_error(_: Request, exc: RequestValidationError):
        return JSONResponse({"detail": "The request was not valid."}, status_code=422)

    @app.exception_handler(Exception)
    async def unexpected(_: Request, exc: Exception):
        log.error("Unexpected error: %s", type(exc).__name__)  # never log data
        return JSONResponse({"detail": "Something went wrong while processing the file."}, status_code=500)

    def run(job: Job, options: Options, overrides: dict) -> dict:
        with job.lock:
            report = process_file(
                job.source, job.kind, job.filename, job.output, settings, options, overrides
            )
        report.update(
            job_id=job.id,
            filename=job.filename,
            file_type=job.kind,
            download_name=job.download_name,
            download_url=f"/api/jobs/{job.id}/download",
            options={"date_order": options.date_order, "output_format": options.output_format},
        )
        return report

    @app.get("/api/config")
    def get_config():
        return {
            "max_upload_mb": settings.max_upload_mb,
            "extensions": list(ALLOWED),
            "job_ttl_minutes": settings.job_ttl_minutes,
        }

    @app.post("/api/jobs")
    def upload(
        file: UploadFile = File(...),
        date_order: str = Form("auto"),
        output_format: str = Form("yyyy-mm-dd"),
    ):
        options = _options(date_order, output_format)
        name = file.filename or "upload"
        ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext not in ALLOWED:
            raise UserError("Unsupported file type. Please upload a .xlsx or .csv file.", 415)
        kind = ALLOWED[ext]
        job = store.create(name, kind)
        try:
            limit = settings.max_upload_mb * 1024 * 1024
            size = 0
            with open(job.source, "wb") as out:
                while chunk := file.file.read(1024 * 1024):
                    size += len(chunk)
                    if size > limit:
                        raise UserError(
                            f"The file is larger than the {settings.max_upload_mb} MB limit.", 413
                        )
                    out.write(chunk)
            if size == 0:
                raise UserError("The file is empty.")
            with open(job.source, "rb") as f:
                head = f.read(4)
            is_zip = head.startswith(b"PK")
            if kind == "xlsx" and not is_zip:
                raise UserError("The file is named .xlsx but is not an Excel workbook.")
            if kind == "csv" and is_zip:
                raise UserError("The file is named .csv but looks like a zipped/Excel file. Upload it as .xlsx.")
            return run(job, options, {})
        except BaseException:
            store.delete(job.id)  # clean up on any failure
            raise

    @app.post("/api/jobs/{job_id}/regenerate")
    def regenerate(job_id: str, body: RegenerateRequest):
        job = store.get(job_id)
        return run(job, _options(body.date_order, body.output_format), _clean_overrides(body.overrides))

    @app.get("/api/jobs/{job_id}/download")
    def download(job_id: str):
        job = store.get(job_id)
        if not job.output.exists():
            raise UserError("The corrected file is not ready.", 404)
        return FileResponse(
            job.output,
            media_type=XLSX_MIME,
            filename=job.download_name,
            headers={"Cache-Control": "no-store"},
        )

    @app.delete("/api/jobs/{job_id}")
    def delete(job_id: str):
        store.delete(job_id)
        return {"ok": True}

    return app


app = create_app()
