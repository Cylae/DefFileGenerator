"""
Unit and integration tests for the Webdyn SunPM tag knowledge engine.
Validates tag inference, legacy normalization, and CSV enrichment according to
SunPM_DEfFile_Tags_v1.xlsx.
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.knowledge_miner import (
    enrich_definition_file,
    enrich_definition_rows,
    infer_sunpm_tag,
    infer_webdyn_tag,
    normalize_tag_to_sunpm,
)
from DefFileGenerator.main import _run_cli


class TestSunPMTags(unittest.TestCase):
    """Test suite for SunPM tag knowledge mining and enrichment."""

    def test_legacy_tag_normalization(self) -> None:
        """Verify legacy camelCase tags normalize to official SunPM standard."""
        test_cases = [
            ("RealPower", "P_AC"),
            ("ActivePowSumkW", "P_AC"),
            ("ActivePow1kW", "P_AC_1"),
            ("ActivePow2kW", "P_AC_2"),
            ("ActivePow3kW", "P_AC_3"),
            ("Pac-Tag", "P_AC"),
            ("EnergyTotal", "E_TOT"),
            ("TotalEnergy", "E_TOT"),
            ("Eac-Tag", "E_TOT"),
            ("EnergyDay", "E_DAY"),
            ("DailyEnergy", "E_DAY"),
            ("EacDay-Tag", "E_DAY"),
            ("ReactivePower", "Q_AC"),
            ("ReactivePowSumkW", "Q_AC"),
            ("ReactivePow1kW", "Q_AC_1"),
            ("GridFrequency", "FREQ"),
            ("GridFrequency1", "FREQ_1"),
            ("CosPhi", "PF"),
            ("cosPHI", "PF"),
            ("NominalPower", "P_RATED"),
            ("DisplayModel", "MODEL"),
            ("DisplaySerialNumber", "SERIAL"),
            ("DisplaySerial", "SERIAL"),
            ("ManufacturerStr", "MANUFACTURER"),
            ("SlaveAddress", "UNIT_ID"),
            ("DeviceAddr", "UNIT_ID"),
            ("cmdOn", "CMD_ON"),
            ("cmdOff", "CMD_OFF"),
            ("cmdPwrPercent", "CMD_P_AC_SET_PERCENT"),
            ("cmdPwr", "CMD_P_AC_SET_ABS"),
            ("WMaxLim_Ena", "CMD_P_ENA"),
            ("WMaxLimPct_RmpTms", "CMD_RAMP_TIME_S"),
            ("WMaxLimPct_RvrtTms", "CMD_REVERT_TIME_S"),
            ("cmdPF", "CMD_PF"),
            ("OutPFSet", "CMD_PF"),
            ("Commit", "CMD_COMMIT"),
            ("Carb-Tag", "AVOIDED_CO2"),
            ("AutoConso-Tag", "SELF_CONSUMPTION_RATIO"),
            ("Conso-Tag", "LOAD_P_AC"),
            ("charge.state", "SOC"),
            ("charge.health", "SOH"),
            ("temperature.battery", "TEMP"),
            ("voltage.battery", "U_DC"),
            ("irradiance", "POA"),
            ("irradiance.horizontal", "GHI"),
        ]
        for legacy, canonical in test_cases:
            self.assertEqual(
                normalize_tag_to_sunpm(legacy),
                canonical,
                f"Expected {legacy} -> {canonical}",
            )
            # Case insensitive check
            self.assertEqual(
                normalize_tag_to_sunpm(legacy.lower()),
                canonical,
                f"Expected {legacy.lower()} -> {canonical}",
            )

    def test_legacy_infer_webdyn_tag_backward_compatibility(self) -> None:
        """Verify legacy infer_webdyn_tag retains exact backward-compatible behavior."""
        self.assertEqual(infer_webdyn_tag("Total Active Power"), "RealPower")
        self.assertEqual(infer_webdyn_tag("Total Forward Active Energy"), "EnergyTotal")
        self.assertEqual(infer_webdyn_tag("Daily Energy Yield"), "EnergyDay")
        self.assertEqual(infer_webdyn_tag("Grid Frequency"), "GridFrequency")
        self.assertEqual(infer_webdyn_tag("Power Factor"), "CosPhi")
        self.assertEqual(infer_webdyn_tag("Phase A Voltage"), "")

    def test_infer_sunpm_ac_voltages(self) -> None:
        """Verify inference of phase-to-phase and phase-to-neutral voltages."""
        # Phase-to-phase
        self.assertEqual(infer_sunpm_tag("Grid AC voltage phase A to phase B"), "U_AC_L1L2")
        self.assertEqual(infer_sunpm_tag("Voltage L1-L2"), "U_AC_L1L2")
        self.assertEqual(infer_sunpm_tag("Tension composee L2-L3"), "U_AC_L2L3")
        self.assertEqual(infer_sunpm_tag("Phase C to Phase A Voltage"), "U_AC_L3L1")

        # Phase-to-neutral
        self.assertEqual(infer_sunpm_tag("Phase A Voltage"), "U_AC_1")
        self.assertEqual(infer_sunpm_tag("Voltage L1-N"), "U_AC_1")
        self.assertEqual(infer_sunpm_tag("Phase B Voltage"), "U_AC_2")
        self.assertEqual(infer_sunpm_tag("Phase C Voltage"), "U_AC_3")

    def test_infer_sunpm_ac_currents_and_powers(self) -> None:
        """Verify inference of per-phase and total AC currents, active, reactive, apparent powers."""
        # Currents
        self.assertEqual(infer_sunpm_tag("Phase A Current"), "I_AC_1")
        self.assertEqual(infer_sunpm_tag("Line 2 Current"), "I_AC_2")
        self.assertEqual(infer_sunpm_tag("Phase C Current"), "I_AC_3")
        self.assertEqual(infer_sunpm_tag("Total AC Current"), "I_AC")

        # Active Power
        self.assertEqual(infer_sunpm_tag("Phase 1 Active Power"), "P_AC_1")
        self.assertEqual(infer_sunpm_tag("Active Power L2"), "P_AC_2")
        self.assertEqual(infer_sunpm_tag("Puissance active L3"), "P_AC_3")
        self.assertEqual(infer_sunpm_tag("Total Active Power"), "P_AC")

        # Reactive & Apparent Power
        self.assertEqual(infer_sunpm_tag("Phase 1 Reactive Power"), "Q_AC_1")
        self.assertEqual(infer_sunpm_tag("Total Reactive Power"), "Q_AC")
        self.assertEqual(infer_sunpm_tag("Phase 2 Apparent Power"), "S_AC_2")
        self.assertEqual(infer_sunpm_tag("Total Apparent Power"), "S_AC")

        # Frequency & Power Factor
        self.assertEqual(infer_sunpm_tag("Grid Frequency"), "FREQ")
        self.assertEqual(infer_sunpm_tag("Total Power Factor"), "PF")

    def test_infer_sunpm_energies_and_dc(self) -> None:
        """Verify inference of daily, total, export, import energies and DC parameters."""
        self.assertEqual(infer_sunpm_tag("Daily Energy Yield"), "E_DAY")
        self.assertEqual(infer_sunpm_tag("Cumulative Total Energy"), "E_TOT")
        self.assertEqual(infer_sunpm_tag("Active Export Energy"), "E_EXP")
        self.assertEqual(infer_sunpm_tag("Active Import Energy"), "E_IMP")

        # DC Bus
        self.assertEqual(infer_sunpm_tag("DC Bus Voltage"), "U_DC")
        self.assertEqual(infer_sunpm_tag("DC Bus Current"), "I_DC")
        self.assertEqual(infer_sunpm_tag("Total DC Power"), "P_DC")

        # MPPT and Strings
        self.assertEqual(infer_sunpm_tag("MPPT 1 Current"), "I_DC_MPPT_1")
        self.assertEqual(infer_sunpm_tag("MPPT 2 Voltage"), "U_DC_MPPT_2")
        self.assertEqual(infer_sunpm_tag("MPPT 3 Power"), "P_DC_MPPT_3")
        self.assertEqual(infer_sunpm_tag("PV 1 Current"), "I_DC_STRING_1")
        self.assertEqual(infer_sunpm_tag("PV String 2 Voltage"), "U_DC_STRING_2")

    def test_infer_sunpm_commands_and_controls(self) -> None:
        """Verify inference of remote inverter control and setting commands."""
        self.assertEqual(infer_sunpm_tag("Start Inverter"), "CMD_ON")
        self.assertEqual(infer_sunpm_tag("Stop Inverter"), "CMD_OFF")
        self.assertEqual(infer_sunpm_tag("Active Power Limitation (%)"), "CMD_P_AC_SET_PERCENT")
        self.assertEqual(infer_sunpm_tag("Active Power Setpoint kW"), "CMD_P_AC_SET_ABS")
        self.assertEqual(infer_sunpm_tag("Power Limitation Enable"), "CMD_P_ENA")
        self.assertEqual(infer_sunpm_tag("Power Ramp Time"), "CMD_RAMP_TIME_S")
        self.assertEqual(infer_sunpm_tag("Power Revert Time"), "CMD_REVERT_TIME_S")
        self.assertEqual(infer_sunpm_tag("Reactive Power Setpoint (%)"), "CMD_Q_AC_SET_PERCENT")
        self.assertEqual(infer_sunpm_tag("Power Factor Setpoint"), "CMD_PF")

    def test_infer_sunpm_sensors_and_batteries(self) -> None:
        """Verify inference of weather sensor and BESS storage tags."""
        # Sensors
        self.assertEqual(infer_sunpm_tag("Plane of Array Irradiance"), "POA")
        self.assertEqual(infer_sunpm_tag("Global Horizontal Irradiance"), "GHI")
        self.assertEqual(infer_sunpm_tag("Ambient Temperature"), "TEMP_AMB")
        self.assertEqual(infer_sunpm_tag("PV Cell Module Temperature"), "TEMP_CELL")
        self.assertEqual(infer_sunpm_tag("Internal Inverter Temperature"), "TEMP_INT")
        self.assertEqual(infer_sunpm_tag("Anemometer Wind Speed"), "WIND_SPEED")

        # Battery
        self.assertEqual(
            infer_sunpm_tag("Battery State of Charge", equipment_type="Battery"), "SOC"
        )
        self.assertEqual(
            infer_sunpm_tag("Battery State of Health", equipment_type="Battery"), "SOH"
        )

    def test_enrich_definition_rows(self) -> None:
        """Verify enriching row dictionaries adds tags and normalizes legacy tags."""
        rows = [
            {"Name": "Total Active Power", "Tag": "RealPower"},
            {"Name": "Phase A Voltage", "Tag": ""},
            {"Name": "Daily Energy", "Tag": "EnergyDay"},
            {"Name": "Custom Unrecognized Value", "Tag": ""},
        ]
        modified = enrich_definition_rows(rows, equipment_type="Inverter")
        self.assertEqual(modified, 3)
        self.assertEqual(rows[0]["Tag"], "P_AC")  # normalized
        self.assertEqual(rows[1]["Tag"], "U_AC_1")  # inferred
        self.assertEqual(rows[2]["Tag"], "E_DAY")  # normalized
        self.assertEqual(rows[3]["Tag"], "")  # unchanged

    def test_enrich_definition_file_roundtrip(self) -> None:
        """Verify reading a definition CSV, enriching tags, and writing atomically."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_in = Path(tmpdir) / "test_def.csv"
            csv_out = Path(tmpdir) / "enriched_def.csv"

            content = [
                ["Modbus", "Inverter", "TestVendor", "TestModel"],
                [
                    "1",
                    "3",
                    "100",
                    "U16",
                    "",
                    "Total Active Power",
                    "RealPower",
                    "1.0",
                    "0.0",
                    "W",
                    "4",
                ],
                ["2", "3", "101", "U16", "", "Phase A Voltage", "", "0.1", "0.0", "V", "4"],
                ["3", "3", "102", "U16", "", "Daily Energy Yield", "", "1.0", "0.0", "kWh", "4"],
            ]
            with open(csv_in, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(content)

            count = enrich_definition_file(str(csv_in), str(csv_out))
            self.assertEqual(count, 3)

            with open(csv_out, encoding="utf-8") as f:
                reader = list(csv.reader(f, delimiter=";"))

            self.assertEqual(reader[1][6], "P_AC")
            self.assertEqual(reader[2][6], "U_AC_1")
            self.assertEqual(reader[3][6], "E_DAY")

    def test_cli_tag_command(self) -> None:
        """Verify running `deffilegen tag` via CLI."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_in = Path(tmpdir) / "cli_test.csv"
            csv_out = Path(tmpdir) / "cli_out.csv"

            content = [
                ["Modbus", "Inverter", "CLI", "Test"],
                [
                    "1",
                    "3",
                    "500",
                    "U16",
                    "",
                    "Grid Frequency",
                    "GridFrequency",
                    "1.0",
                    "0.0",
                    "Hz",
                    "4",
                ],
                ["2", "3", "501", "U16", "", "Start Inverter", "", "1.0", "0.0", "", "1"],
            ]
            with open(csv_in, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(content)

            _run_cli(["tag", str(csv_in), "-o", str(csv_out)])

            with open(csv_out, encoding="utf-8") as f:
                rows = list(csv.reader(f, delimiter=";"))

            self.assertEqual(rows[1][6], "FREQ")
            self.assertEqual(rows[2][6], "CMD_ON")


if __name__ == "__main__":
    unittest.main()
