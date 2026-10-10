"""
Targeted unit tests (Wave 23) for edge branches in def_gen.py, main.py, and web/app.py.
"""

import io
import os
import tempfile
from unittest.mock import AsyncMock, patch

import pytest
from fastapi import HTTPException
from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    CSVHeaderConfig,
    Generator,
    GeneratorConfig,
    generate_template,
    run_generator,
)
from web.app import _save_upload, app


def test_validate_address_raw_invalid_len():
    # RAW type with non-positive byte length
    assert Generator.validate_address("40001_0", "RAW") is False
    assert Generator.validate_address("40001_-4", "RAW") is False
    # RAW type with odd byte length
    assert Generator.validate_address("40001_3", "RAW") is False


def test_validate_address_bits_invalid_slice():
    # BITS type with start_bit < 0 or bit_length <= 0 or start + len > 16
    assert Generator.validate_address("40001_-1_8", "BITS") is False
    assert Generator.validate_address("40001_0_0", "BITS") is False
    assert Generator.validate_address("40001_10_8", "BITS") is False


def test_validate_address_exception_handling():
    # Trigger exception in validate_address during split or normalization
    assert Generator.validate_address("invalid_address_parts_x_y_z", "U16") is False


def test_parse_numeric_fraction_edge_cases():
    # Invalid fraction formats
    assert Generator._parse_numeric("1/2/3", default=0.0) == 0.0
    assert Generator._parse_numeric("1/0", default=0.0) == 0.0
    assert Generator._parse_numeric("abc/def", default=0.0) == 0.0


def test_write_output_csv_validation_failure_staging():
    # write_output_csv with validation_strict=True when output is a stream and generator yields invalid data
    gen = Generator()
    stream = io.StringIO()
    invalid_rows = [
        {
            "Info1": "INVALID_INFO1",
            "Info2": "99999",
            "Info3": "INVALID_TYPE",
            "Info4": "",
            "Name": "test",
            "Tag": "test",
            "CoefA": "1.0",
            "CoefB": "0.0",
            "Unit": "",
            "Action": "4",
        }
    ]
    hdr = CSVHeaderConfig(manufacturer="Mfg", model="Mod")
    # Should fail validation and return False
    res = gen.write_output_csv(stream, invalid_rows, config=hdr, validation_strict=True)
    assert res is False


def test_write_output_csv_non_oserror_reraised():
    # When output is not stream staging and an unexpected non-OSError exception occurs, it should be re-raised
    gen = Generator()
    with patch("csv.writer") as mock_writer:
        mock_writer.side_effect = RuntimeError("Unexpected internal error")
        with pytest.raises(RuntimeError, match="Unexpected internal error"):
            gen.write_output_csv("non_existent_output.csv", [], config=CSVHeaderConfig())


def test_generate_template_oserror():
    # Force OSError when writing template
    with patch("DefFileGenerator.def_gen.staged_text_output", side_effect=OSError("Disk full")):
        assert generate_template("some_file.csv", mode="input") is False


def test_run_generator_missing_input_file():
    config = GeneratorConfig(input_file="non_existent_file_12345.csv", output="out.csv")
    assert run_generator(config) is False


def test_web_health_endpoint():
    client = TestClient(app)
    res = client.get("/api/health")
    assert res.status_code == 200


@pytest.mark.anyio
async def test_save_upload_oversized():
    mock_file = AsyncMock()
    # Mock file.read returning chunks total > MAX_UPLOAD_BYTES
    chunk_10mb = b"x" * (10 * 1024 * 1024 + 1)
    mock_file.read.side_effect = [chunk_10mb, b""]
    mock_file.close = AsyncMock()

    with tempfile.TemporaryDirectory() as tmpdir:
        dest = os.path.join(tmpdir, "test.bin")
        with pytest.raises(HTTPException) as exc_info:
            await _save_upload(mock_file, dest)
        assert exc_info.value.status_code == 413
