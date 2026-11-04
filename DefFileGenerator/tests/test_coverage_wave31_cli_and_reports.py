"""Tests for Wave 31: CLI subcommands, report generation, batch edge cases, and auto-fix integration."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import pytest

from DefFileGenerator.main import (
    audit_command,
    batch_command,
    generate_command,
    report_command,
    run_command,
)
from DefFileGenerator.report_generator import (
    analyze_definition_file,
    analyze_definitions_directory,
    export_report_json,
    generate_directory_technical_report,
)


def test_audit_command_nonexistent_target():
    args = argparse.Namespace(target="nonexistent_file_or_dir_12345.csv")
    with pytest.raises(SystemExit) as exc:
        audit_command(args)
    assert exc.value.code == 1


def test_audit_command_single_file_with_fix_and_json(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
):
    csv_file = tmp_path / "modbus_test.csv"
    # Create CSV with duplicate rows and non-canonical tag
    csv_file.write_text(
        "modbusRTU;Inverter;Vendor;Model;;;;;;;\n"
        "1;3;100;U16;;Active Power;ACTIVE_POWER;1;0;W;4\n"
        "1;3;100;U16;;Active Power;ACTIVE_POWER;1;0;W;4\n",
        encoding="utf-8",
    )
    output_fixed = tmp_path / "fixed_output.csv"
    args = argparse.Namespace(
        target=str(csv_file),
        output=str(output_fixed),
        fix=True,
        json=True,
        strict=False,
        strict_overlap=True,
        verbose=True,
    )
    audit_command(args)
    captured = capsys.readouterr()
    data = json.loads(captured.out)
    assert "is_valid" in data
    assert "fix_applied" in data
    assert data["fix_applied"]["duplicates_removed"] == 1
    assert output_fixed.exists()


def test_audit_command_single_file_human_readable_logging(tmp_path: Path):
    csv_file = tmp_path / "modbus_bad.csv"
    # Create file with invalid row syntax that causes audit errors
    csv_file.write_text(
        "modbusRTU;Inverter;Vendor;Model;;;;;;;\n1;3;NOT_AN_INT;U16;;Bad Reg;P_AC;1;0;W;4\n",
        encoding="utf-8",
    )
    args = argparse.Namespace(
        target=str(csv_file),
        fix=False,
        json=False,
        strict=True,
        strict_overlap=True,
        verbose=True,
    )
    with pytest.raises(SystemExit) as exc:
        audit_command(args)
    assert exc.value.code == 1


def test_audit_command_directory_mode_json_and_fix(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
):
    sub_dir = tmp_path / "defs"
    sub_dir.mkdir()
    f1 = sub_dir / "valid_def.csv"
    f1.write_text(
        "modbusRTU;Inverter;Vendor;M1;;;;;;;\n1;3;100;U16;;Active Power;P_AC;1;0;W;4\n",
        encoding="utf-8",
    )
    f2 = sub_dir / "dup_def.csv"
    f2.write_text(
        "modbusRTU;Inverter;Vendor;M2;;;;;;;\n"
        "1;3;200;U16;;Voltage;U_AC_1;1;0;V;4\n"
        "1;3;200;U16;;Voltage;U_AC_1;1;0;V;4\n",
        encoding="utf-8",
    )

    args = argparse.Namespace(
        target=str(sub_dir),
        fix=True,
        json=True,
        strict=False,
        strict_overlap=False,
    )
    audit_command(args)
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert payload["summary"]["total_files"] == 2
    assert "fix_applied" in payload
    assert payload["fix_applied"]["duplicates_removed"] == 1


def test_audit_command_directory_mode_human_readable_errors(tmp_path: Path):
    sub_dir = tmp_path / "defs_err"
    sub_dir.mkdir()
    f1 = sub_dir / "broken.csv"
    f1.write_text(
        "modbusRTU;Inverter;Vendor;M1;;;;;;;\n1;3;INVALID;U16;;Bad;P_AC;1;0;W;4\n",
        encoding="utf-8",
    )
    args = argparse.Namespace(
        target=str(sub_dir),
        fix=False,
        json=False,
        strict=True,
        strict_overlap=True,
    )
    with pytest.raises(SystemExit) as exc:
        audit_command(args)
    assert exc.value.code == 1


def test_generate_command_with_report(tmp_path: Path):
    input_csv = tmp_path / "clean_registers.csv"
    input_csv.write_text(
        "Address,Name,Type,Factor,Offset,Unit\n"
        "100,Active Power,U16,1,0,W\n"
        "101,Reactive Power,S16,1,0,VAr\n",
        encoding="utf-8",
    )
    output_def = tmp_path / "generated_definition.csv"
    report_file = tmp_path / "tech_report.md"

    args = argparse.Namespace(
        input_file=str(input_csv),
        output=str(output_def),
        manufacturer="TestMfg",
        model="TestModel",
        protocol="modbusRTU",
        category="Inverter",
        forced_write="",
        offset=0,
        report=str(report_file),
        template=False,
    )
    generate_command(args)
    assert output_def.exists()
    assert report_file.exists()
    assert "TestMfg" in report_file.read_text(encoding="utf-8")


def test_run_command_validations_and_report(tmp_path: Path):
    # Test missing manufacturer/model
    args_bad = argparse.Namespace(template=False, manufacturer=None, model=None, input_file=None)
    with pytest.raises(SystemExit) as exc:
        run_command(args_bad)
    assert exc.value.code == 1

    # Test missing input_file
    args_bad2 = argparse.Namespace(template=False, manufacturer="M", model="X", input_file=None)
    with pytest.raises(SystemExit) as exc:
        run_command(args_bad2)
    assert exc.value.code == 1

    # Test non-existent input_file
    args_bad3 = argparse.Namespace(
        template=False, manufacturer="M", model="X", input_file="missing_path.csv"
    )
    with pytest.raises(SystemExit) as exc:
        run_command(args_bad3)
    assert exc.value.code == 1

    # Test successful run with --report
    input_csv = tmp_path / "valid_raw.csv"
    input_csv.write_text(
        "Address,Name,Type,Factor,Offset,Unit\n500,Energy Total,U32,1,0,kWh\n",
        encoding="utf-8",
    )
    out_csv = tmp_path / "run_out.csv"
    rep_md = tmp_path / "run_report.md"

    args_ok = argparse.Namespace(
        template=False,
        manufacturer="SMA",
        model="SunnyTripower",
        input_file=str(input_csv),
        output=str(out_csv),
        report=str(rep_md),
        protocol="modbusRTU",
        category="Inverter",
        forced_write="",
        no_validate=False,
        force=True,
    )
    run_command(args_ok)
    assert out_csv.exists()
    assert rep_md.exists()


def test_report_command_file_and_directory(tmp_path: Path):
    # Nonexistent path raises sys.exit(1)
    with pytest.raises(SystemExit) as exc:
        report_command(argparse.Namespace(target="does_not_exist_xyz.csv"))
    assert exc.value.code == 1

    # Single file report
    def_file = tmp_path / "test_def.csv"
    def_file.write_text(
        "modbusRTU;Inverter;Vendor;Mod1;;;;;;;\n1;3;100;U16;;Power;P_AC;1;0;W;4\n",
        encoding="utf-8",
    )
    report_out = tmp_path / "single_report.md"
    report_command(argparse.Namespace(target=str(def_file), output=str(report_out), json=None))
    assert report_out.exists()

    # Directory report with JSON export
    defs_dir = tmp_path / "defs_rep"
    defs_dir.mkdir()
    (defs_dir / "d1.csv").write_text(
        "modbusRTU;Meter;Vendor;M1;;;;;;;\n1;3;10;U32;;Energy;E_TOT;1;0;kWh;4\n",
        encoding="utf-8",
    )
    dir_report_out = tmp_path / "dir_report.md"
    dir_json_out = tmp_path / "dir_metrics.json"
    report_command(
        argparse.Namespace(
            target=str(defs_dir),
            output=str(dir_report_out),
            json=str(dir_json_out),
        )
    )
    assert dir_report_out.exists()
    assert dir_json_out.exists()
    metrics = json.loads(dir_json_out.read_text(encoding="utf-8"))
    assert metrics["total_files"] == 1


def test_batch_command_skip_existing_and_exceptions(tmp_path: Path):
    in_dir = tmp_path / "batch_input"
    in_dir.mkdir()
    out_dir = tmp_path / "batch_output"
    out_dir.mkdir()

    # Create dummy supported file
    csv_doc = in_dir / "Huawei_SUN2000.csv"
    csv_doc.write_text("Address,Name,Type\n100,Power,U16\n", encoding="utf-8")

    # Pre-create output file to trigger skipping without --force
    expected_out = out_dir / "huawei_sun2000_definition.csv"
    expected_out.write_text("existing content", encoding="utf-8")

    args = argparse.Namespace(
        input_dir=str(in_dir),
        output_dir=str(out_dir),
        tag=False,
        strict=False,
        force=False,
        report=None,
        protocol="modbusRTU",
        manufacturer=None,
        category=None,
    )
    # Since only 1 file is skipped, success_count=0 and fail_count=0, so no exit(1)
    batch_command(args)
    # Content must NOT be overwritten
    assert expected_out.read_text(encoding="utf-8") == "existing content"

    # Now create an empty file that triggers fail_count > 0 and exit(1) when success_count=0
    empty_csv = in_dir / "Empty_Meter.csv"
    empty_csv.write_text("JustHeaderWithoutRegisters\n", encoding="utf-8")
    with pytest.raises(SystemExit) as exc:
        batch_command(args)
    assert exc.value.code == 1


def test_report_generator_edge_actions_and_invalid_file(tmp_path: Path):
    # Test action=9, other action (e.g. 5), short row (<7 items), and invalid definition
    csv_file = tmp_path / "custom_actions.csv"
    csv_file.write_text(
        "modbusRTU;Sensor;Kipp;CMP11;;;;;;;\n"
        "1;3;10;U16;;Alarm;ALARM_STATUS;1;0;;9\n"
        "1;3;20;U16;;Custom;TEMP_MOD;1;0;;5\n"
        "short;row\n"
        "1;3;INVALID_ADDR;U16;;Broken;U_AC_1;1;0;;4\n",
        encoding="utf-8",
    )

    stats = analyze_definition_file(csv_file)
    assert stats.action_9_count == 1
    assert stats.other_action_count == 1
    assert not stats.is_valid

    # Directory report covering invalid files and other actions
    dir_stats = analyze_definitions_directory(tmp_path)
    assert dir_stats.invalid_files >= 1
    assert "Other Actions" in dir_stats.action_counts

    # Test file write in generate_directory_technical_report
    rep_file = tmp_path / "dir_tech.md"
    out_content = generate_directory_technical_report(tmp_path, output_report_path=rep_file)
    assert rep_file.exists()
    assert len(out_content) > 100

    # Test JSON export
    json_path = tmp_path / "metrics_test.json"
    export_report_json(dir_stats, json_path)
    assert json_path.exists()
