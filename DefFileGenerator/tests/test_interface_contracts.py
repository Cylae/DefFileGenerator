"""Regression checks for generated-output integrity across the public interfaces."""

import asyncio
import csv
import io
import json
import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch

import httpx
import pytest
import starlette.formparsers as formparsers
from fastapi.testclient import TestClient

import DefFileGenerator.gui as gui_module
import web.app as web_app
from DefFileGenerator.def_gen import ValidationReport


def test_upload_content_length_rejected_before_parsing():
    with (
        patch.object(web_app, "MAX_UPLOAD_BYTES", 32),
        patch.object(web_app, "MAX_MULTIPART_OVERHEAD_BYTES", 0),
        patch.object(web_app, "_save_upload") as save,
        TestClient(web_app.app) as client,
    ):
        response = client.post("/api/convert", files={"file": ("input.csv", b"X" * 128)})
    assert response.status_code == 413
    save.assert_not_called()


def test_chunked_upload_rejected_and_partial_spool_closed():
    created = []
    original = formparsers.SpooledTemporaryFile

    def tracked_spool(*args, **kwargs):
        spool = original(*args, **kwargs)
        created.append(spool)
        return spool

    async def exercise():
        body = (
            b'--test\r\nContent-Disposition: form-data; name="file"; filename="input.csv"\r\n'
            b"Content-Type: text/csv\r\n\r\n" + b"X" * 512 + b"\r\n--test--\r\n"
        )

        async def chunks():
            for start in range(0, len(body), 64):
                yield body[start : start + 64]

        async with httpx.AsyncClient(
            transport=httpx.ASGITransport(app=web_app.app), base_url="http://test"
        ) as client:
            return await client.post(
                "/api/convert",
                content=chunks(),
                headers={"Content-Type": "multipart/form-data; boundary=test"},
            )

    with (
        patch.object(web_app, "MAX_UPLOAD_BYTES", 200),
        patch.object(web_app, "MAX_MULTIPART_OVERHEAD_BYTES", 0),
        patch.object(formparsers, "SpooledTemporaryFile", side_effect=tracked_spool),
        patch.object(web_app, "_save_upload") as save,
    ):
        response = asyncio.run(exercise())
    assert response.status_code == 413
    save.assert_not_called()
    assert created and all(spool.closed for spool in created)


def test_extra_uploaded_files_rejected_before_conversion():
    with patch.object(web_app, "_save_upload") as save, TestClient(web_app.app) as client:
        response = client.post(
            "/api/convert",
            files=[
                ("file", ("one.csv", b"Address,Name,Type\n100,N,U16\n")),
                ("extra", ("two.csv", b"X")),
            ],
        )
    assert response.status_code == 413
    save.assert_not_called()


@pytest.mark.parametrize("file_output", [False, True])
def test_intermediate_csv_escapes_formulas_and_preserves_signed_numbers(
    tmp_path, capsys, file_output
):
    from DefFileGenerator.main import main

    source = tmp_path / "input.csv"
    source.write_text("Address,Name,Type,Factor\n100,=1+1,U16,-0.25\n", encoding="utf-8")
    args = ["extract", str(source)]
    output = tmp_path / "output.csv"
    if file_output:
        args.extend(["-o", str(output)])
    main(args)
    content = output.read_text(encoding="utf-8") if file_output else capsys.readouterr().out
    row = next(csv.DictReader(io.StringIO(content)))
    assert row["Name"] == "'=1+1"
    assert row["Factor"] == "-0.25"


def test_conversion_does_not_publish_overlapping_definition():
    with TestClient(web_app.app) as client:
        response = client.post(
            "/api/convert",
            files={"file": ("overlap.csv", b"Address,Name,Type\n100,Wide,U32\n101,Narrow,U16\n")},
        )
    assert response.status_code == 400


def test_conversion_does_not_publish_empty_definition_after_all_rows_are_skipped():
    with TestClient(web_app.app) as client:
        response = client.post(
            "/api/convert",
            files={"file": ("map.csv", b"Address,Name,Type\n100,N,UNSUPPORTED_TYPE\n")},
        )
    assert response.status_code == 400


def test_conversion_preview_and_count_describe_downloaded_rows():
    source = (
        b"Address,Name,Type,Factor,ScaleFactor\n"
        b"100,Good,U16,0.1,-1\n"
        b"101,Invalid,UNSUPPORTED_TYPE,1,0\n"
        b"102,Good,U16,1,0\n"
    )
    with TestClient(web_app.app) as client:
        response = client.post("/api/convert", files={"file": ("map.csv", source)})
        assert response.status_code == 200
        result = response.json()
        rows = list(csv.reader(io.StringIO(result["csv_content"]), delimiter=";"))[1:]
        assert result["register_count"] == len(rows) == 2
        assert result["extracted_count"] == 3
        assert result["preview"][0]["Factor"] == rows[0][7] == "0.010000"
        assert [row["Tag"] for row in result["preview"]] == [row[6] for row in rows]
        assert [row["Name"] for row in result["preview"]] == [row[5] for row in rows]
        validation = client.post(
            "/api/validate", files={"file": ("result.csv", result["csv_content"].encode())}
        ).json()
        assert result["is_valid"] == validation["valid"]


def test_web_register_limit_stops_consuming_at_first_excess_row():
    consumed = []

    def mapped_rows():
        for address in range(10):
            consumed.append(address)
            yield {"Address": str(address), "Name": f"N{address}", "Type": "U16"}

    with (
        patch.object(web_app, "MAX_WEB_REGISTERS", 2),
        patch.object(web_app.Extractor, "map_and_clean", return_value=mapped_rows()),
        TestClient(web_app.app) as client,
    ):
        response = client.post(
            "/api/convert", files={"file": ("map.csv", b"Address,Name,Type\n100,N,U16\n")}
        )
    assert response.status_code == 400
    assert consumed == [0, 1, 2]


@pytest.mark.parametrize("endpoint", ["convert", "validate"])
def test_health_request_can_run_while_document_worker_is_busy(endpoint):
    started = threading.Event()
    release = threading.Event()
    finished = threading.Event()
    owner = web_app if endpoint == "convert" else web_app.Generator
    method = "run_generator" if endpoint == "convert" else "validate_csv_detailed"
    original = getattr(owner, method)

    def blocked_worker(*args, **kwargs):
        started.set()
        release.wait(2)
        try:
            return original(*args, **kwargs)
        finally:
            finished.set()

    async def exercise():
        transport = httpx.ASGITransport(app=web_app.app)
        async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
            source = (
                b"Address,Name,Type\n100,N,U16\n"
                if endpoint == "convert"
                else b"modbusRTU;Inverter;M;X;;;;;;;\n1;3;100;U16;;N;n;1;0;;4\n"
            )
            request = asyncio.create_task(
                client.post(f"/api/{endpoint}", files={"file": ("input.csv", source)})
            )
            try:
                assert await asyncio.to_thread(started.wait, 3)
                health = await asyncio.wait_for(client.get("/api/health"), timeout=1)
                assert health.status_code == 200
                assert not finished.is_set(), "Document processing blocked the request event loop"
            finally:
                release.set()
                await request

    with patch.object(owner, method, new=blocked_worker):
        asyncio.run(exercise())


def test_cli_generation_reports_output_write_failure(tmp_path):
    source = tmp_path / "input.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "DefFileGenerator.main",
            "generate",
            str(source),
            "--manufacturer",
            "M",
            "--model",
            "X",
            "-o",
            str(tmp_path / "missing" / "out.csv"),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1


def test_cli_generate_rejects_overlaps_before_writing_stdout(tmp_path):
    source = tmp_path / "overlap.csv"
    source.write_text("Address,Name,Type\n100,Wide,U32\n101,Narrow,U16\n", encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "DefFileGenerator.main",
            "generate",
            str(source),
            "--manufacturer",
            "M",
            "--model",
            "X",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1
    assert result.stdout == ""


@pytest.mark.parametrize("command", ["generate", "run"])
def test_cli_generation_failure_cannot_be_hidden_by_old_valid_output(tmp_path, command):
    from DefFileGenerator.main import main

    source = tmp_path / "input.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    target = tmp_path / "output.csv"
    previous = "modbusRTU;Inverter;Old;X;;;;;;;\n1;3;100;U16;;Old;old;1;0;;4\n"
    target.write_text(previous, encoding="utf-8")
    with (
        patch("DefFileGenerator.main.run_generator", return_value=False),
        pytest.raises(SystemExit) as error,
    ):
        main(
            [
                command,
                str(source),
                "--manufacturer",
                "M",
                "--model",
                "X",
                "-o",
                str(target),
                "--force",
            ]
        )
    assert error.value.code == 1
    assert target.read_text(encoding="utf-8") == previous


@pytest.mark.parametrize("placement", ["before", "after"])
def test_cli_quiet_flag_works_on_either_side_of_subcommand(tmp_path, placement):
    source = tmp_path / "input.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    args = ["generate", str(source), "--manufacturer", "M", "--model", "X"]
    args.insert(0, "-q") if placement == "before" else args.append("-q")
    result = subprocess.run(
        [sys.executable, "-m", "DefFileGenerator.main", *args],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0
    assert "INFO" not in result.stderr


def _headless_app(source: Path, output: Path):
    app = gui_module.DefFileGenApp.__new__(gui_module.DefFileGenApp)
    app.is_processing = False
    for name, value in {
        "entry_input_file": str(source),
        "entry_output_file": str(output),
        "entry_mfg": "M",
        "entry_model": "X",
        "entry_offset": "0",
        "opt_protocol": "modbusRTU",
        "opt_category": "Inverter",
    }.items():
        widget = MagicMock()
        widget.get.return_value = value
        setattr(app, name, widget)
    app.btn_convert = MagicMock()
    app.lbl_results_badge = MagicMock()
    app.tree_preview = MagicMock()
    return app


@pytest.mark.parametrize("alias", [False, True])
def test_gui_rejects_overwriting_source_even_through_hardlink(tmp_path, alias):
    source = tmp_path / "map.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    output = tmp_path / "alias.csv" if alias else source
    if alias:
        os.link(source, output)
    app = _headless_app(source, output)
    with (
        patch.object(gui_module.messagebox, "showerror") as error,
        patch.object(gui_module.threading, "Thread") as worker,
    ):
        app._start_conversion_thread()
    worker.assert_not_called()
    error.assert_called_once()
    assert not app.is_processing
    assert source.read_text(encoding="utf-8").startswith("Address,Name,Type")


def test_gui_requires_overwrite_confirmation(tmp_path):
    source = tmp_path / "map.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    output = tmp_path / "out.csv"
    output.write_text("SENTINEL", encoding="utf-8")
    app = _headless_app(source, output)
    with (
        patch.object(gui_module.messagebox, "askyesno", return_value=False) as confirm,
        patch.object(gui_module.threading, "Thread") as worker,
    ):
        app._start_conversion_thread()
    confirm.assert_called_once()
    worker.assert_not_called()
    assert output.read_text(encoding="utf-8") == "SENTINEL"


def test_gui_core_failure_does_not_report_success_for_old_output(tmp_path):
    source = tmp_path / "input.csv"
    source.write_text("Address,Name,Type\n100,N,U16\n", encoding="utf-8")
    target = tmp_path / "output.csv"
    previous = "modbusRTU;Inverter;Old;X;;;;;;;\n1;3;100;U16;;Old;old;1;0;;4\n"
    target.write_text(previous, encoding="utf-8")
    app = _headless_app(source, target)
    app.root = MagicMock()
    with patch.object(gui_module, "run_generator", return_value=False):
        app._run_conversion_worker(str(source), str(target), "M", "X", "modbusRTU", "Inverter", 0)
    assert app.root.after.call_args.args[1] == app._on_conversion_error
    assert target.read_text(encoding="utf-8") == previous


def test_gui_json_report_export_is_json(tmp_path):
    output = tmp_path / "report.json"
    app = gui_module.DefFileGenApp.__new__(gui_module.DefFileGenApp)
    app.last_validation_report = ValidationReport(
        is_valid=True, register_count=1, issues=[], stats={"errors": 0, "warnings": 0}
    )
    with (
        patch.object(gui_module.filedialog, "asksaveasfilename", return_value=str(output)),
        patch.object(gui_module.messagebox, "showinfo"),
    ):
        app._export_validation_report()
    report = json.loads(output.read_text(encoding="utf-8"))
    assert report["is_valid"] is True
    assert report["register_count"] == 1


@pytest.mark.parametrize("command", ["extract", "generate", "run"])
@pytest.mark.parametrize("alias", [False, True])
def test_cli_never_overwrites_source_even_with_force(tmp_path, command, alias):
    from DefFileGenerator.main import main

    source = tmp_path / "input.csv"
    original = "Address,Name,Type\n100,N,U16\n"
    source.write_text(original, encoding="utf-8")
    destination = tmp_path / "alias.csv" if alias else source
    if alias:
        os.link(source, destination)
    args = [command, str(source), "-o", str(destination), "--force"]
    if command != "extract":
        args.extend(["--manufacturer", "M", "--model", "X"])
    with pytest.raises(SystemExit) as error:
        main(args)
    assert error.value.code == 1
    assert source.read_text(encoding="utf-8") == original


def test_late_extraction_failure_preserves_existing_intermediate_output(tmp_path):
    from DefFileGenerator.main import main

    source = tmp_path / "source.csv"
    source.write_text(
        "Address,Name,Type\n100,Good,U16\n101," + "N" * (csv.field_size_limit() + 1) + ",U16\n",
        encoding="utf-8",
    )
    target = tmp_path / "intermediate.csv"
    target.write_text("SENTINEL", encoding="utf-8")
    with pytest.raises(SystemExit) as error:
        main(["extract", str(source), "-o", str(target), "--force"])
    assert error.value.code == 1
    assert target.read_text(encoding="utf-8") == "SENTINEL"
    assert set(tmp_path.iterdir()) == {source, target}


def test_frontend_validity_download_and_stale_request_contracts():
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js is needed to execute the browser script contract tests")
    tests = Path(__file__).parent
    result = subprocess.run(
        [
            node,
            str(tests / "interface_frontend_contracts.js"),
            str(tests.parent.parent / "web" / "static" / "app.js"),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
