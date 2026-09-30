"""Regression coverage for Modbus spans, byte ordering, and definition decoding."""

import csv
import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import Generator, GeneratorConfig, generate_template, run_generator


@pytest.mark.parametrize(
    ("dtype", "last_valid", "first_invalid"),
    [
        ("U32", "65534", "65535"),
        ("F32_WB", "65534", "65535"),
        ("I64_W", "65532", "65533"),
        ("MAC", "65533", "65534"),
        ("IPV6", "65528", "65529"),
        ("STRING", "65534_3", "65535_3"),
        ("RAW", "65534_4", "65535_4"),
    ],
)
def test_entire_register_span_must_fit_modbus_address_space(dtype, last_valid, first_invalid):
    assert Generator.validate_address(last_valid, dtype)
    assert not Generator.validate_address(first_invalid, dtype)
    assert Generator.validate_address(first_invalid, dtype, strict=False)


def test_huge_string_length_is_rejected_without_float_overflow():
    address = "0_" + "9" * 400
    assert not Generator.validate_address(address, "STRING")
    assert Generator.get_register_count("STRING", address) == (int("9" * 400) + 1) // 2


def test_processed_rows_reject_spans_shifted_out_of_range():
    rows = [{"Name": "Power", "Address": "65534", "Type": "U32"}]
    assert list(Generator().process_rows(rows, address_offset=1)) == []


def test_strict_definition_validation_rejects_out_of_range_span(tmp_path):
    path = tmp_path / "definition.csv"
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(["modbusRTU", "Inverter", "M", "X", "", "", "", "", "", "", ""])
        writer.writerow(["1", "3", "65535", "U32", "", "Power", "power", "1", "0", "W", "4"])

    report = Generator().validate_csv_detailed(str(path), strict=True)
    assert not report.is_valid
    assert any(issue.code == "INVALID_ADDRESS" for issue in report.issues)


@pytest.mark.parametrize("payload", [b"modbusRTU;Inverter\n\xff", b"\xff\xfe\x00"])
def test_malformed_definition_encoding_returns_validation_issue(tmp_path, payload):
    path = tmp_path / "invalid-encoding.csv"
    path.write_bytes(payload)

    report = Generator().validate_csv_detailed(str(path), strict=True)
    assert not report.is_valid
    assert report.stats["errors"] == 1
    assert report.issues[0].code == "IO_ERROR"


@pytest.mark.parametrize(
    ("raw_type", "expected"),
    [
        ("float32 word swap", "F32_W"),
        ("float32 word-swap", "F32_W"),
        ("signed integer 64 word swap", "I64_W"),
        ("float32 byte swap", "F32_B"),
        ("uint32 byte-swap", "U32_B"),
        ("float32 word swap and byte swap", "F32_WB"),
    ],
)
def test_explicit_swap_aliases_preserve_requested_byte_order(raw_type, expected):
    assert Generator.normalize_type(raw_type) == expected
    assert Generator.validate_type(expected)


def _processed_row(address="100", tag="power"):
    return {
        "Info1": "3",
        "Info2": address,
        "Info3": "U16",
        "Info4": "",
        "Name": "Power",
        "Tag": tag,
        "CoefA": "1.000000",
        "CoefB": "0.000000",
        "Unit": "W",
        "Action": "4",
    }


def test_writer_reports_failed_replacement_and_preserves_target(tmp_path, monkeypatch):
    path = tmp_path / "definition.csv"
    path.write_bytes(b"previous definition")

    def failed_replace(*args):
        raise OSError("target is locked")

    monkeypatch.setattr("DefFileGenerator.def_gen.os.replace", failed_replace)
    assert Generator.write_output_csv(str(path), [_processed_row()], "M", "X") is False
    assert path.read_bytes() == b"previous definition"
    assert list(tmp_path.iterdir()) == [path]


def test_writer_returns_success_after_validating_staged_definition(tmp_path):
    path = tmp_path / "definition.csv"
    assert (
        Generator.write_output_csv(str(path), [_processed_row()], "M", "X", validation_strict=True)
        is True
    )
    assert Generator().validate_csv(str(path), strict=True)


def test_strict_writer_preserves_target_when_staged_definition_is_invalid(tmp_path):
    path = tmp_path / "definition.csv"
    path.write_bytes(b"previous definition")
    rows = [_processed_row("100"), _processed_row("101")]

    assert Generator.write_output_csv(str(path), rows, "M", "X", validation_strict=True) is False
    assert path.read_bytes() == b"previous definition"
    assert list(tmp_path.iterdir()) == [path]


def test_generator_reports_deferred_input_failure_without_replacing_target(tmp_path):
    path = tmp_path / "definition.csv"
    path.write_bytes(b"previous definition")

    def rows():
        yield {"Name": "Power", "Address": "100", "Type": "U16"}
        raise RuntimeError("document parser failed")

    config = GeneratorConfig(output=str(path), manufacturer="M", model="X")
    assert run_generator(config, input_data=rows()) is False
    assert path.read_bytes() == b"previous definition"
    assert list(tmp_path.iterdir()) == [path]


def test_generator_forwards_strict_validation_before_replacement(tmp_path):
    path = tmp_path / "definition.csv"
    path.write_bytes(b"previous definition")
    rows = [
        {"Name": "Power", "Address": "100", "Type": "U32"},
        {"Name": "Voltage", "Address": "101", "Type": "U16"},
    ]
    config = GeneratorConfig(output=str(path), strict_validation=True)

    assert run_generator(config, input_data=rows) is False
    assert path.read_bytes() == b"previous definition"
    assert list(tmp_path.iterdir()) == [path]


def test_template_reports_write_failure(tmp_path):
    path = tmp_path / "missing-parent" / "template.csv"
    assert generate_template(str(path)) is False
    assert run_generator(GeneratorConfig(output=str(path), template=True)) is False


def test_generator_reports_missing_input():
    assert run_generator(GeneratorConfig()) is False


@pytest.mark.parametrize("strict_validation", [True, False])
def test_configured_generator_rejects_all_unsupported_rows_without_replacing_target(
    tmp_path, strict_validation
):
    path = tmp_path / "definition.csv"
    path.write_bytes(b"previous definition")
    config = GeneratorConfig(output=str(path), strict_validation=strict_validation)
    rows = [{"Name": "Power", "Address": "100", "Type": "unsupported vendor type"}]

    assert run_generator(config, input_data=rows) is False
    assert path.read_bytes() == b"previous definition"
    assert list(tmp_path.iterdir()) == [path]


def test_unconfigured_writer_preserves_empty_template_compatibility(tmp_path):
    path = tmp_path / "definition.csv"
    assert Generator.write_output_csv(str(path), [], "M", "X") is True
    assert Generator().validate_csv(str(path), strict=True)


@pytest.mark.parametrize("use_input_data", [True, False])
def test_generator_rejects_source_as_output_for_library_callers(tmp_path, use_input_data):
    source = tmp_path / "source.csv"
    original = b"Address,Name,Type\n100,Power,U16\n"
    source.write_bytes(original)
    config = GeneratorConfig(input_file=str(source), output=str(source))
    rows = [{"Address": "100", "Name": "Power", "Type": "U16"}] if use_input_data else None

    assert run_generator(config, input_data=rows) is False
    assert source.read_bytes() == original


def test_generator_rejects_hard_link_to_source(tmp_path):
    source = tmp_path / "source.csv"
    destination = tmp_path / "definition.csv"
    original = b"Address,Name,Type\n100,Power,U16\n"
    source.write_bytes(original)
    os.link(source, destination)
    config = GeneratorConfig(input_file=str(source), output=str(destination))

    assert run_generator(config) is False
    assert source.read_bytes() == destination.read_bytes() == original


def test_failed_template_write_preserves_previous_destination(tmp_path, monkeypatch):
    path = tmp_path / "template.csv"
    path.write_bytes(b"previous template")
    original_writer = csv.writer

    class FailingWriter:
        def __init__(self, stream, **kwargs):
            self.writer = original_writer(stream, **kwargs)

        def writerow(self, row):
            self.writer.writerow(row)

        def writerows(self, rows):
            self.writer.writerow(rows[0])
            raise csv.Error("template serialization failed")

    monkeypatch.setattr("DefFileGenerator.def_gen.csv.writer", FailingWriter)

    assert generate_template(str(path)) is False
    assert path.read_bytes() == b"previous template"
    assert list(tmp_path.iterdir()) == [path]


def test_strict_validator_uses_normalized_function_code_for_overlap(tmp_path):
    path = tmp_path / "definition.csv"
    path.write_text(
        "modbusRTU;Inverter;M;X;;;;;;;\n"
        "1;3;100;U16;;First;first;1;0;V;4\n"
        "2; 3 ;100;U16;;Second;second;1;0;V;4\n",
        encoding="utf-8",
    )

    report = Generator().validate_csv_detailed(str(path), strict=True)
    assert not report.is_valid
    assert any(issue.code == "ADDRESS_OVERLAP" for issue in report.issues)


@pytest.mark.parametrize("template", [True, False])
def test_standalone_generator_script_exits_with_failure_on_write_error(tmp_path, template):
    source = tmp_path / "source.csv"
    source.write_text("Address,Name,Type\n100,Power,U16\n", encoding="utf-8")
    script = Path(__file__).resolve().parents[1] / "def_gen.py"
    arguments = ["--template"] if template else [str(source)]

    result = subprocess.run(
        [sys.executable, str(script), *arguments, "-o", str(tmp_path / "missing" / "out.csv")],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode == 1
    assert "Error" in result.stderr


def test_configured_stdout_matches_successful_unconfigured_output(monkeypatch):
    rows = [{"Name": "Power; measured", "Address": "100", "Type": "U16", "Unit": "°C"}]
    expected = io.StringIO()
    monkeypatch.setattr("sys.stdout", expected)
    assert run_generator(GeneratorConfig(manufacturer="M", model="X"), input_data=rows)
    output = io.StringIO()
    monkeypatch.setattr("sys.stdout", output)

    assert run_generator(
        GeneratorConfig(manufacturer="M", model="X", strict_validation=True), input_data=rows
    )
    assert output.getvalue() == expected.getvalue()
    assert not output.getvalue().startswith("\ufeff")
    assert not output.closed


def test_configured_stdout_rejects_overlaps_without_emitting_csv(monkeypatch):
    output = io.StringIO()
    monkeypatch.setattr("sys.stdout", output)
    rows = [
        {"Name": "Power", "Address": "100", "Type": "U32"},
        {"Name": "Voltage", "Address": "101", "Type": "U16"},
    ]

    assert run_generator(GeneratorConfig(strict_validation=True), input_data=rows) is False
    assert output.getvalue() == ""
    assert not output.closed


def test_configured_stream_rejects_deferred_iterator_failure_without_emitting_csv():
    output = io.StringIO("previous stream content")
    output.seek(0, io.SEEK_END)

    def rows():
        yield _processed_row()
        raise RuntimeError("document parser failed")

    assert Generator.write_output_csv(output, rows(), "M", "X", validation_strict=True) is False
    assert output.getvalue() == "previous stream content"
    assert not output.closed


@pytest.mark.parametrize("strict_validation", [True, False])
def test_configured_stream_rejects_empty_conversion_without_emitting_csv(strict_validation):
    output = io.StringIO()
    assert Generator.write_output_csv(output, [], validation_strict=strict_validation) is False
    assert output.getvalue() == ""
    assert not output.closed


def test_configured_stream_copy_is_bounded_and_retains_caller_ownership():
    class RecordingStream(io.StringIO):
        def __init__(self):
            super().__init__()
            self.write_sizes = []

        def write(self, value):
            self.write_sizes.append(len(value))
            return super().write(value)

    output = RecordingStream()
    rows = [_processed_row(str(address), f"power_{address}") for address in range(2000)]

    assert Generator.write_output_csv(output, rows, "M", "X", validation_strict=True) is True
    assert len(output.write_sizes) >= 2
    assert max(output.write_sizes) <= 65536
    assert len(list(csv.reader(io.StringIO(output.getvalue()), delimiter=";"))) == 2001
    assert not output.closed


@pytest.mark.parametrize("result_kind", ["success", "invalid", "deferred_failure"])
def test_configured_stream_closes_and_removes_owned_staging_files(
    tmp_path, monkeypatch, result_kind
):
    output = io.StringIO()
    original_temporary_file = tempfile.NamedTemporaryFile
    staging_files = []

    def temporary_file(**kwargs):
        staging_file = original_temporary_file(dir=tmp_path, **kwargs)
        staging_files.append(staging_file)
        return staging_file

    def failed_rows():
        yield _processed_row()
        raise RuntimeError("document parser failed")

    monkeypatch.setattr("DefFileGenerator.def_gen.tempfile.NamedTemporaryFile", temporary_file)
    if result_kind == "success":
        rows = [_processed_row()]
    elif result_kind == "invalid":
        rows = [_processed_row(), _processed_row("101")]
    else:
        rows = failed_rows()

    result = Generator.write_output_csv(output, rows, "M", "X", validation_strict=True)

    assert result is (result_kind == "success")
    assert len(staging_files) == 1
    assert staging_files[0].closed
    assert not Path(staging_files[0].name).exists()
    assert not output.closed
