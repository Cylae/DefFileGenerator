#!/usr/bin/env python3
"""
Unit and integration tests for 'deffilegen batch' CLI subcommand and metadata inference.
"""

from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

import pytest

from DefFileGenerator.main import (
    batch_command,
    infer_metadata_from_filename,
    main,
)


class TestMetadataInference:
    """Tests filename-based manufacturer, model, and category inference."""

    def test_infer_huawei_inverter(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("Huawei_SUN2000_50KTL.pdf")
        assert mfg == "Huawei"
        assert "SUN2000" in model
        assert cat == "Inverter"

    def test_infer_sma_inverter(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("SMA_STP5000TL.xlsx")
        assert mfg == "SMA"
        assert "STP5000TL" in model
        assert cat == "Inverter"

    def test_infer_abb_meter(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("ABB_M4M_Meter_Registers.xlsx")
        assert mfg == "ABB"
        assert cat == "Meter"

    def test_infer_sensor(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("Hukseflux_Pyranometer_SR20.csv")
        assert mfg == "Hukseflux"
        assert cat == "Sensor"

    def test_infer_battery(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("BESS_Battery_Rack_Cell_Telemetry.csv")
        assert cat == "Battery"

    def test_infer_communication_logger(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("Huawei_SmartLogger_SL3000.csv")
        assert mfg == "Huawei"
        assert cat == "Communication"

    def test_infer_io_module(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("Webdyn_Sens_IO_16DI.csv")
        assert mfg == "Webdyn"
        assert cat == "Io"

    def test_infer_generic_fallback(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("unlabeled_document.pdf")
        assert mfg == "Manufacturer"
        assert model == "unlabeled-document"
        assert cat == "Inverter"

    def test_infer_empty_stem(self) -> None:
        mfg, model, cat = infer_metadata_from_filename("___")
        assert mfg == "Manufacturer"
        assert model == "Model"
        assert cat == "Inverter"


class TestBatchCommand:
    """Integration tests for batch processing command."""

    def test_batch_missing_directory(self) -> None:
        args = argparse.Namespace(input_dir="non_existent_directory_12345")
        with pytest.raises(SystemExit) as exc:
            batch_command(args)
        assert exc.value.code == 1

    def test_batch_empty_directory(self) -> None:
        with tempfile.TemporaryDirectory() as empty_dir:
            args = argparse.Namespace(
                input_dir=empty_dir,
                output_dir=os.path.join(empty_dir, "out"),
                protocol="modbusRTU",
                category=None,
                manufacturer=None,
                tag=False,
                strict=False,
                force=False,
                report=None,
            )
            with pytest.raises(SystemExit) as exc:
                batch_command(args)
            assert exc.value.code == 1

    def test_batch_conversion_with_sample_excel(self) -> None:
        sample_file = Path("sample_inverter_registers.xlsx")
        assert sample_file.exists()

        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(input_dir)

            dest_sample = os.path.join(input_dir, "SampleBrand_ModelX_inverter.xlsx")
            with open(sample_file, "rb") as src, open(dest_sample, "wb") as dst:
                dst.write(src.read())

            report_file = os.path.join(temp_dir, "batch_report.md")

            args = argparse.Namespace(
                input_dir=input_dir,
                output_dir=output_dir,
                protocol="modbusRTU",
                category="Inverter",
                manufacturer="SampleBrand",
                tag=True,
                strict=True,
                force=True,
                report=report_file,
            )

            batch_command(args)

            output_files = list(Path(output_dir).glob("*.csv"))
            assert len(output_files) == 1
            csv_path = output_files[0]
            content = csv_path.read_text(encoding="utf-8-sig")
            assert "modbusRTU;Inverter;SampleBrand;" in content

            assert os.path.exists(report_file)
            report_content = Path(report_file).read_text(encoding="utf-8")
            assert (
                "Rapport Technique" in report_content
                or "Technical Report" in report_content
                or "SampleBrand" in report_content
            )

    def test_batch_cli_dispatch(self) -> None:
        sample_file = Path("sample_inverter_registers.xlsx")
        assert sample_file.exists()

        with tempfile.TemporaryDirectory() as temp_dir:
            out_dir = os.path.join(temp_dir, "cli_out")
            main(
                [
                    "batch",
                    str(sample_file),
                    "-o",
                    out_dir,
                    "--force",
                ]
            )
            output_files = list(Path(out_dir).glob("*.csv"))
            assert len(output_files) >= 1

    def test_batch_skip_existing_without_force(self) -> None:
        sample_file = Path("sample_inverter_registers.xlsx")
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(input_dir)
            os.makedirs(output_dir)

            dest_sample = os.path.join(input_dir, "SampleBrand_ModelX_inverter.xlsx")
            with open(sample_file, "rb") as src, open(dest_sample, "wb") as dst:
                dst.write(src.read())

            # Pre-create the output file
            existing_out = os.path.join(output_dir, "samplebrand_modelx_inverter_definition.csv")
            with open(existing_out, "w") as f:
                f.write("existing content")

            args = argparse.Namespace(
                input_dir=input_dir,
                output_dir=output_dir,
                protocol="modbusRTU",
                category="Inverter",
                manufacturer="SampleBrand",
                tag=False,
                strict=False,
                force=False,
                report=None,
            )
            batch_command(args)

            # Check that existing content was not overwritten
            assert Path(existing_out).read_text() == "existing content"

    def test_batch_skip_unsupported_extensions(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(input_dir)

            dummy_txt = os.path.join(input_dir, "notes.txt")
            Path(dummy_txt).write_text("not supported")

            args = argparse.Namespace(
                input_dir=input_dir,
                output_dir=output_dir,
                protocol="modbusRTU",
                category=None,
                manufacturer=None,
                tag=False,
                strict=False,
                force=False,
                report=None,
            )
            with pytest.raises(SystemExit) as exc:
                batch_command(args)
            assert exc.value.code == 1
