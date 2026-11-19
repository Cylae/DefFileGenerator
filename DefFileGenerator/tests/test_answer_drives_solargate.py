"""
Test suite for Answer Drives Solargate Inverter Modbus extraction and WebdynSunPM definition.
"""

import os
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.main import infer_metadata_from_filename

SOLARGATE_PDF_PATH = r"c:\Users\alexandre.calais\Downloads\ADSolargateManENHWSW 02-11-09 (1).pdf"
SOLARGATE_DEF_PATH = Path("Definitions/INV_AnswerDrives_Solargate.csv")


class TestAnswerDrivesSolargate:
    """Validates Answer Drives Solargate Modbus extraction and definition integrity."""

    def test_metadata_inference(self) -> None:
        """Asserts that filename patterns containing 'solargate' or 'answer drives' infer correct metadata."""
        mfg, model, cat = infer_metadata_from_filename(
            "c:/docs/ADSolargateManENHWSW 02-11-09 (1).pdf"
        )
        assert mfg == "Answer Drives"
        assert model == "Solargate"
        assert cat == "Inverter"

        mfg2, model2, cat2 = infer_metadata_from_filename("INV_AnswerDrives_Solargate.csv")
        assert mfg2 == "Answer Drives"
        assert model2 == "Solargate"
        assert cat2 == "Inverter"

    def test_generated_definition_validity(self) -> None:
        """Verifies that Definitions/INV_AnswerDrives_Solargate.csv is valid according to Webdyn standards."""
        assert SOLARGATE_DEF_PATH.exists(), f"Missing definition file: {SOLARGATE_DEF_PATH}"
        gen = Generator()
        report = gen.validate_csv_detailed(str(SOLARGATE_DEF_PATH))
        assert report.is_valid, f"Validation failed with issues: {report.issues}"
        assert report.register_count >= 280, (
            f"Expected at least 280 registers, found {report.register_count}"
        )

    def test_definition_telemetry_registers(self) -> None:
        """Ensures key telemetry and supervisory registers are correctly mapped with addresses and types."""
        assert SOLARGATE_DEF_PATH.exists()
        with open(SOLARGATE_DEF_PATH, encoding="utf-8-sig") as f:
            lines = [line.strip().split(";") for line in f if line.strip()]

        header_line = lines[0]
        assert header_line[0] == "Modbus"
        assert header_line[1] == "Inverter"
        assert header_line[2] == "Answer Drives"
        assert header_line[3] == "Solargate"

        records_by_addr = {parts[2]: parts for parts in lines[1:] if len(parts) >= 11}

        # DC Bus Voltage at 40206
        assert "40206" in records_by_addr
        r_vdc = records_by_addr["40206"]
        assert r_vdc[1] == "3"  # Holding Register
        assert r_vdc[3] in ("I16", "U16")
        assert "DCbusvoltage" in r_vdc[5]
        assert r_vdc[10] == "4"  # Read Action

        # RMS AC Current at 40207
        assert "40207" in records_by_addr
        r_curr = records_by_addr["40207"]
        assert "RMSAcLinecurrent" in r_curr[5]

        # Power Factor at 40211
        assert "40211" in records_by_addr
        r_pf = records_by_addr["40211"]
        assert "PowerFactor" in r_pf[5]

        # AC Line Voltage at 40213
        assert "40213" in records_by_addr
        r_vac = records_by_addr["40213"]
        assert "ACLineVoltage" in r_vac[5]

        # Frequency at 40214
        assert "40214" in records_by_addr
        r_freq = records_by_addr["40214"]
        assert "frequency" in r_freq[5].lower()

        # Active Power at 40217
        assert "40217" in records_by_addr
        r_p = records_by_addr["40217"]
        assert "ActivePower" in r_p[5]

        # String register length syntax (e.g. 40342_2)
        assert "40342_2" in records_by_addr
        r_str = records_by_addr["40342_2"]
        assert r_str[3] == "STRING"
        assert "Drivesize" in r_str[5]

    @pytest.mark.skipif(
        not os.path.exists(SOLARGATE_PDF_PATH),
        reason="Answer Drives Solargate PDF document not available on local machine",
    )
    def test_live_pdf_extraction(self) -> None:
        """Performs live end-to-end PDF extraction when the source document is available."""
        extractor = Extractor()
        tables = list(extractor.extract_from_pdf(SOLARGATE_PDF_PATH, pages="160-173"))
        assert len(tables) >= 20, f"Expected at least 20 tables, found {len(tables)}"

        mapped = list(extractor.map_and_clean(tables))
        assert len(mapped) >= 280, f"Expected at least 280 registers, found {len(mapped)}"

        # Check telemetry presence
        addrs = {r["Address"] for r in mapped}
        assert "40206" in addrs
        assert "40207" in addrs
        assert "40211" in addrs
        assert "40213" in addrs
        assert "40214" in addrs
        assert "40217" in addrs
        assert "40342_2" in addrs

    def test_tag_inference_telemetry_accuracy(self) -> None:
        """Verifies that SunPM telemetry tags are accurately inferred for key measurements."""
        from DefFileGenerator.knowledge_miner import infer_sunpm_tag

        assert infer_sunpm_tag("DCbusvoltage(Vdc)", equipment_type="Inverter") == "U_DC"
        assert infer_sunpm_tag("RMSAcLinecurrent(Is)", equipment_type="Inverter") == "I_AC"
        assert infer_sunpm_tag("ACLinePowerFactor(cosφ)", equipment_type="Inverter") == "PF"
        assert infer_sunpm_tag("ACLinefrequency", equipment_type="Inverter") == "FREQ"
        assert infer_sunpm_tag("ACLineReactivePower", equipment_type="Inverter") == "Q_AC"
        assert infer_sunpm_tag("ActivePower", equipment_type="Inverter") == "P_AC"
        assert infer_sunpm_tag("Dailyenergy", equipment_type="Inverter") == "E_DAY"
        assert infer_sunpm_tag("Totalenergy(lowword)", equipment_type="Inverter") == "E_TOT"

    def test_reserved_and_coils_do_not_leak_analog_tags(self) -> None:
        """Guards that placeholder/reserved rows or status/trip coils never receive false-positive analog tags."""
        from DefFileGenerator.knowledge_miner import infer_sunpm_tag

        assert infer_sunpm_tag("Reserved", equipment_type="Inverter") == ""
        assert infer_sunpm_tag("Reserved_1", equipment_type="Inverter") == ""
        assert infer_sunpm_tag("NotUsed", equipment_type="Inverter") == ""
        assert infer_sunpm_tag("Nu", equipment_type="Inverter") == ""
        assert infer_sunpm_tag("-", equipment_type="Inverter") == ""

    def test_register_type_resolution_robustness(self) -> None:
        """Verifies that function codes, written forms, and split tokens resolve cleanly to Modbus types."""
        gen = Generator()
        assert gen._determine_info1("0x03-ReadHoldingRegister") == "3"
        assert gen._determine_info1("0x10WriteMultipleHolding") == "3"
        assert gen._determine_info1("Register") == "3"
        assert gen._determine_info1("Holding Register") == "3"
        assert gen._determine_info1("Coil") == "1"

    @pytest.mark.skipif(
        not os.path.exists(SOLARGATE_PDF_PATH),
        reason="Answer Drives Solargate PDF document not available on local machine",
    )
    def test_cli_run_end_to_end(self, tmp_path: Path) -> None:
        """Ensures the CLI run command can extract, generate, and validate directly from the PDF."""
        from DefFileGenerator.main import main

        out_csv = tmp_path / "INV_AnswerDrives_Solargate_CLI.csv"
        test_args = [
            "run",
            SOLARGATE_PDF_PATH,
            "--pages",
            "160-173",
            "--manufacturer",
            "Answer Drives",
            "--model",
            "Solargate",
            "--output",
            str(out_csv),
            "--force",
        ]
        main(test_args)
        assert out_csv.exists()

        gen = Generator()
        report = gen.validate_csv_detailed(str(out_csv))
        assert report.is_valid
        assert report.register_count >= 280
