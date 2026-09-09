#!/usr/bin/env python3
"""
FastAPI Web Backend for WebdynSunPM DefFileGenerator.

Provides REST API endpoints for converting register documentation files
and validating definition files, completely decoupled from the core generator.
"""

import os
import sys
import tempfile
from typing import Any

# Inject parent directory to import core modules cleanly
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402
from fastapi.responses import FileResponse, JSONResponse  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator  # noqa: E402
from DefFileGenerator.extractor import Extractor, peek_generator  # noqa: E402

app = FastAPI(
    title="WebdynSunPM Definition Generator API",
    description="REST API for parsing Modbus register documentation and generating validated WebdynSunPM definition files.",
    version="0.2.1",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


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

    ext = os.path.splitext(file.filename)[1].lower()
    allowed_exts = {".pdf", ".xlsx", ".xlsm", ".xltx", ".xltm", ".csv", ".xml"}
    if ext not in allowed_exts:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format '{ext}'. Allowed: PDF, Excel, CSV, XML.",
        )

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, file.filename)
        output_path = os.path.join(temp_dir, "generated_definition.csv")

        # Save uploaded bytes to temp file
        contents = await file.read()
        with open(input_path, "wb") as f:
            f.write(contents)

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
        except Exception as e:
            raise HTTPException(
                status_code=400, detail=f"Failed to extract registers from file: {e}"
            ) from e

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

        # Collect preview rows (up to 500)
        preview_rows = []
        full_mapped = list(mapped_peeked)
        for r in full_mapped[:500]:
            preview_rows.append(r)

        config = GeneratorConfig(
            input_file=input_path,
            output=output_path,
            manufacturer=manufacturer,
            model=model,
            protocol=protocol,
            category=category,
            forced_write=forced_write,
            address_offset=0,
        )

        run_generator(config, input_data=full_mapped)

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="Failed to generate output CSV definition.")

        generator = Generator()
        is_valid = generator.validate_csv(output_path, strict=False)

        # Read generated CSV content
        with open(output_path, encoding="utf-8-sig") as f:
            csv_content = f.read()

        filename_mfg = "".join(
            c for c in manufacturer.lower().replace(" ", "_") if c.isalnum() or c == "_"
        )
        filename_model = "".join(
            c for c in model.lower().replace(" ", "_") if c.isalnum() or c == "_"
        )
        out_filename = f"{filename_mfg}_{filename_model}_definition.csv"

        return JSONResponse(
            content={
                "success": True,
                "is_valid": is_valid,
                "filename": out_filename,
                "register_count": len(full_mapped),
                "preview": preview_rows,
                "csv_content": csv_content,
            }
        )


@app.post("/api/validate")
async def validate_file(file: UploadFile = File(...)) -> Any:
    """Validates an uploaded WebdynSunPM CSV definition file."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Uploaded file must have a filename.")

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, file.filename)
        contents = await file.read()
        with open(input_path, "wb") as f:
            f.write(contents)

        generator = Generator()
        is_valid = generator.validate_csv(input_path, strict=True)
        return JSONResponse(content={"filename": file.filename, "valid": is_valid})


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

    uvicorn.run(app, host="0.0.0.0", port=8000)
