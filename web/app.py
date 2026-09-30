#!/usr/bin/env python3
"""
FastAPI Web Backend for WebdynSunPM DefFileGenerator.

Provides REST API endpoints for converting register documentation files
and validating definition files, completely decoupled from the core generator.
"""

import logging
import os
import sys
import tempfile
from itertools import islice
from typing import Any

# Inject parent directory to import core modules cleanly
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile  # noqa: E402
from fastapi.concurrency import run_in_threadpool  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402
from fastapi.responses import FileResponse, JSONResponse  # noqa: E402
from fastapi.routing import APIRoute  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402
from starlette.exceptions import HTTPException as StarletteHTTPException  # noqa: E402
from starlette.formparsers import MultiPartException  # noqa: E402
from starlette.types import ASGIApp, Message, Receive, Scope, Send  # noqa: E402

from DefFileGenerator.def_gen import (  # noqa: E402
    Generator,
    GeneratorConfig,
    ValidationReport,
    run_generator,
)
from DefFileGenerator.extractor import Extractor, peek_generator  # noqa: E402
from DefFileGenerator.io_utils import definition_preview  # noqa: E402

logger = logging.getLogger("DefFileGenerator.web")

MAX_UPLOAD_BYTES = 10 * 1024 * 1024
UPLOAD_CHUNK_BYTES = 1024 * 1024
MAX_WEB_REGISTERS = 65536
MAX_MULTIPART_OVERHEAD_BYTES = 64 * 1024
UPLOAD_PATHS = frozenset({"/api/convert", "/api/validate"})


class UploadBodyLimitMiddleware:
    """Bound bytes before multipart parsing can spool upload files."""

    def __init__(self, app: ASGIApp):
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if (
            scope["type"] != "http"
            or scope["method"] != "POST"
            or scope["path"] not in UPLOAD_PATHS
        ):
            await self.app(scope, receive, send)
            return

        maximum = MAX_UPLOAD_BYTES + MAX_MULTIPART_OVERHEAD_BYTES
        rejection = JSONResponse(
            status_code=413,
            content={
                "detail": f"Upload request exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MiB limit plus bounded multipart metadata."
            },
        )
        for name, value in scope.get("headers", []):
            if name.lower() == b"content-length":
                try:
                    too_large = int(value) > maximum
                except ValueError:
                    too_large = (
                        False  # The actual-byte limit also covers malformed or absent lengths.
                    )
                if too_large:
                    await rejection(scope, receive, send)
                    return

        total = 0
        exceeded = False

        async def bounded_receive() -> Message:
            nonlocal total, exceeded
            message = await receive()
            if message["type"] == "http.request":
                total += len(message.get("body", b""))
                if total > maximum:
                    exceeded = True
                    # This exception is recognized by all supported Starlette multipart
                    # parsers, which close partial spool files before propagating it.
                    raise MultiPartException("Upload request body exceeds its maximum size.")
            return message

        async def bounded_send(message: Message) -> None:
            if not exceeded:
                await send(message)

        try:
            await self.app(scope, bounded_receive, bounded_send)
        except MultiPartException:
            if not exceeded:
                raise
        if exceeded:
            await rejection(scope, receive, send)


class UploadRoute(APIRoute):
    """Apply file-count limits before FastAPI parses the same cached form."""

    def get_route_handler(self):
        original_handler = super().get_route_handler()

        async def handler(request: Request):
            if request.method == "POST" and request.url.path in UPLOAD_PATHS:
                try:
                    await request.form(max_files=1, max_fields=6)
                except StarletteHTTPException as exc:
                    if "Too many files" in str(exc.detail):
                        raise HTTPException(
                            status_code=413, detail="Only one uploaded file is allowed."
                        ) from None
                    raise
            return await original_handler(request)

        return handler


app = FastAPI(
    title="WebdynSunPM Definition Generator API",
    description="REST API for parsing Modbus register documentation and generating validated WebdynSunPM definition files.",
    version="0.2.1",
)
app.router.route_class = UploadRoute
app.add_middleware(UploadBodyLimitMiddleware)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


async def _save_upload(file: UploadFile, destination: str) -> None:
    """Persist an upload in bounded chunks and reject oversized payloads."""
    total = 0
    try:
        with open(destination, "wb") as output:
            while True:
                chunk = await file.read(UPLOAD_CHUNK_BYTES)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail=f"Uploaded file exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MiB limit.",
                    )
                output.write(chunk)
    finally:
        await file.close()


@app.get("/api/health")
def health_check() -> dict[str, str]:
    return {"status": "ok", "version": "0.2.1"}


@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    manufacturer: str = Form("Manufacturer"),
    model: str = Form("Model"),
    protocol: str = Form("modbusRTU"),
    category: str = Form("Inverter"),
    address_offset: int = Form(0),
    forced_write: str = Form(""),
) -> Any:
    """Converts an uploaded documentation file (PDF, Excel, CSV, XML) to a WebdynSunPM definition CSV."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Uploaded file must have a filename.")

    # Sanitize filename to prevent path traversal attack
    safe_filename = os.path.basename(file.filename.replace("\\", "/").strip())
    if not safe_filename or safe_filename in (".", ".."):
        raise HTTPException(status_code=400, detail="Invalid filename provided.")

    ext = os.path.splitext(safe_filename)[1].lower()
    allowed_exts = {".pdf", ".xlsx", ".xlsm", ".xltx", ".xltm", ".csv", ".xml"}
    if ext not in allowed_exts:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format '{ext}'. Allowed: PDF, Excel, CSV, XML.",
        )

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, f"source_input{ext}")
        output_path = os.path.join(temp_dir, "generated_definition.csv")

        await _save_upload(file, input_path)
        result = await run_in_threadpool(
            _convert_document,
            input_path,
            output_path,
            ext,
            safe_filename,
            manufacturer,
            model,
            protocol,
            category,
            address_offset,
            forced_write,
        )
        return JSONResponse(content=result)


def _convert_document(
    input_path: str,
    output_path: str,
    ext: str,
    safe_filename: str,
    manufacturer: str,
    model: str,
    protocol: str,
    category: str,
    address_offset: int,
    forced_write: str,
) -> dict[str, Any]:
    """Process the complete document outside the request event loop."""
    extractor = Extractor()
    try:
        if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
            raw_data = extractor.extract_from_excel(input_path)
        elif ext == ".pdf":
            raw_data = extractor.extract_from_pdf(input_path)
        elif ext == ".csv":
            raw_data = extractor.extract_from_csv(input_path)
        elif ext == ".xml":
            raw_data = extractor.extract_from_xml(input_path)
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format.")
    except HTTPException:
        raise
    except Exception:
        logger.exception("Extraction error for %s", safe_filename)
        raise HTTPException(
            status_code=400, detail="Failed to extract registers from uploaded file."
        ) from None

    try:
        has_data, raw_data_peeked = peek_generator(raw_data)
        if not has_data:
            raise HTTPException(
                status_code=400, detail="No readable register tables found in uploaded file."
            )

        mapped_gen = extractor.map_and_clean(raw_data_peeked, address_offset)
        has_regs, mapped_peeked = peek_generator(mapped_gen)
        if not has_regs:
            raise HTTPException(
                status_code=400, detail="No valid registers mapped after field cleaning step."
            )

        full_mapped = list(islice(mapped_peeked, MAX_WEB_REGISTERS + 1))
        if len(full_mapped) > MAX_WEB_REGISTERS:
            raise HTTPException(
                status_code=400,
                detail=f"Uploaded file contains more than {MAX_WEB_REGISTERS} registers, exceeding the maximum allowable limit of {MAX_WEB_REGISTERS}.",
            )

        config = GeneratorConfig(
            input_file=input_path,
            output=output_path,
            manufacturer=manufacturer,
            model=model,
            protocol=protocol,
            category=category,
            forced_write=forced_write,
            address_offset=0,
            strict_validation=True,
        )
        if run_generator(config, input_data=full_mapped) is False:
            raise HTTPException(
                status_code=400, detail="Core generator processing failed for uploaded file."
            )
    except HTTPException:
        raise
    except Exception:
        logger.exception("Core engine processing failed for %s", safe_filename)
        raise HTTPException(
            status_code=400, detail="Core generator processing failed for uploaded file."
        ) from None

    if not os.path.exists(output_path):
        raise HTTPException(status_code=500, detail="Failed to generate output CSV definition.")

    try:
        report = Generator().validate_csv_detailed(output_path, strict=True)
        if not report.is_valid:
            raise HTTPException(
                status_code=400, detail="Generated definition failed strict validation."
            )
        with open(output_path, encoding="utf-8-sig") as generated:
            csv_content = generated.read()
    except HTTPException:
        raise
    except Exception:
        logger.exception("Validation check encountered an error on generated CSV")
        raise HTTPException(
            status_code=400, detail="Failed to validate generated definition CSV."
        ) from None

    preview_rows = definition_preview(csv_content, limit=500)

    mfg_clean = "".join(
        c for c in manufacturer.lower().replace(" ", "_") if c.isalnum() or c == "_"
    )[:50]
    filename_mfg = mfg_clean or "manufacturer"
    model_clean = "".join(c for c in model.lower().replace(" ", "_") if c.isalnum() or c == "_")[
        :50
    ]
    filename_model = model_clean or "model"
    out_filename = f"{filename_mfg}_{filename_model}_definition.csv"

    return {
        "success": True,
        "is_valid": report.is_valid,
        "filename": out_filename,
        "register_count": report.register_count,
        "extracted_count": len(full_mapped),
        "preview": preview_rows,
        "csv_content": csv_content,
        "issues": _serialize_issues(report),
        "stats": report.stats,
    }


def _serialize_issues(report: ValidationReport) -> list[dict[str, Any]]:
    return [
        {
            "line": issue.line,
            "severity": issue.severity,
            "code": issue.code,
            "field": issue.field,
            "message": issue.message,
        }
        for issue in report.issues[:1000]
    ]


def _validate_document(input_path: str, safe_filename: str) -> dict[str, Any]:
    try:
        report = Generator().validate_csv_detailed(input_path, strict=True)
    except Exception:
        logger.exception("Validation error for %s", safe_filename)
        raise HTTPException(
            status_code=400, detail="Failed to validate uploaded definition CSV."
        ) from None
    return {
        "filename": safe_filename,
        "valid": report.is_valid,
        "register_count": report.register_count,
        "issues": _serialize_issues(report),
        "stats": report.stats,
    }


@app.post("/api/validate")
async def validate_file(file: UploadFile = File(...)) -> Any:
    """Validates an uploaded WebdynSunPM CSV definition file."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Uploaded file must have a filename.")

    safe_filename = os.path.basename(file.filename.replace("\\", "/").strip())
    if not safe_filename or safe_filename in (".", ".."):
        raise HTTPException(status_code=400, detail="Invalid filename provided.")

    ext = os.path.splitext(safe_filename)[1].lower()
    if ext not in {".csv", ".txt"}:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format '{ext}'. Validation requires a CSV definition file (.csv, .txt).",
        )

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, f"source_input{ext}")
        await _save_upload(file, input_path)

        result = await run_in_threadpool(_validate_document, input_path, safe_filename)
        return JSONResponse(content=result)


# Serve static web frontend assets
static_dir = os.path.join(os.path.dirname(__file__), "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/")
def read_root():
    index_path = os.path.join(static_dir, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "WebdynSunPM API server running. Frontend assets not found."}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
