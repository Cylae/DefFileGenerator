"""
Adversarial Verification Suite for SMA Complete Product Ecosystem:
Residential (Sunny Boy, Sunny Boy Storage), Commercial (Sunny Tripower, Sunny Tripower X),
Decentralized C&I (Core 1, Core 2), Hybrid Storage (Sunny Tripower Smart Energy, Sunny Island),
Utility Scale (Sunny Central, SC Storage, SC2xxx), Communication Gateways (Data Manager M, Cluster Controller,
Com Gateway, WebBox), Smart Meters (Energy Meter, Home Manager), and SunSpec 3-Phase.

Validates:
1. All 27 SMA definition files exist in production/ and Definitions/SMA/
2. Exact 5 Apollo-Standard Pillars compliance for every single file
3. Canonical SunPM tag uniqueness (no duplicate single-tags within a file)
4. Strict sequential 1..N index matching row number
5. Valid Modbus function codes (1, 2, 3, 4) and certified Webdyn action codes (0, 1, 4, 9, 10)
6. 6-decimal coefficient formatting (e.g. 1.000000, 0.001000)
7. Byte-identical synchronization across production, Definitions/SMA, Definitions, generated_definitions, and tagged_definitions
8. Domain-specific telemetry coverage (MPPT trackers, battery BMS, 3-phase grid AC, plant controls)
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PRODUCTION_DIR = REPO_ROOT / "production"
DEFINITIONS_SMA_DIR = REPO_ROOT / "Definitions" / "SMA"
DEFINITIONS_DIR = REPO_ROOT / "Definitions"
GENERATED_DIR = REPO_ROOT / "generated_definitions"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"

EXPECTED_SMA_FILES = [
    # Residential
    "INV_SMA_SunnyBoy.csv",
    "INV_SMA_SunnyBoyStorage.csv",
    # Commercial Three-Phase
    "INV_SMA_SunnyTripower.csv",
    "INV_SMA_SunnyTripower_X.csv",
    "INV_SMA_STP_X.csv",
    "INV_SMA_Core1_Core2.csv",
    "INV_SMA_Core1.csv",
    "INV_SMA_Core2.csv",
    "INV_SMA_STPCORE1_50-41.csv",
    "INV_SMA_STP_110-60.csv",
    "INV_SMA_STP_125-70.csv",
    # Hybrid & Energy Storage
    "INV_SMA_SunnyTripower_SmartEnergy.csv",
    "INV_SMA_STP_5-10-SE40.csv",
    "INV_SMA_SunnyIsland.csv",
    # Utility Scale Central
    "INV_SMA_SunnyCentral.csv",
    "INV_SMA_SUNNY_CENTRAL.csv",
    "INV_SMA_SC2xxx.csv",
    "INV_SMA_SunnyCentral_Storage.csv",
    # Communication Gateways & Plant Controllers
    "COM_SMA_DataManager_M.csv",
    "COM_SMA_ClusterController.csv",
    "COM_SMA_ComGateway.csv",
    "INV_SMA_SC_Webbox.csv",
    # Smart Energy Meters & Grid Analyzers
    "MET_SMA_EnergyMeter.csv",
    "MET_SMA_HomeManager.csv",
    # SunSpec Standard Inverters & Interoperable Profiles
    "INV_SMA_SunSpec_Inverter.csv",
    "INV_SMA_Sunspec.csv",
    "INV_SMA_Generic.csv",
    "INV_SMA_Modbus.csv",
]


class TestSmaDefinitionsSuite(unittest.TestCase):
    """Test suite certifying all 28 SMA ecosystem definitions."""

    def test_all_files_exist_in_all_target_directories(self) -> None:
        """Ensure all 28 files exist in production, Definitions/SMA, Definitions, generated_definitions, and tagged_definitions."""
        for fname in EXPECTED_SMA_FILES:
            for d in (
                PRODUCTION_DIR,
                DEFINITIONS_SMA_DIR,
                DEFINITIONS_DIR,
                GENERATED_DIR,
                TAGGED_DIR,
            ):
                fpath = d / fname
                self.assertTrue(fpath.is_file(), f"Missing {fname} in {d.name}")

    def test_byte_identical_redundant_synchronization(self) -> None:
        """Verify byte-identical synchronization across redundant directories."""
        for fname in EXPECTED_SMA_FILES:
            prod_bytes = (PRODUCTION_DIR / fname).read_bytes()
            for d in (DEFINITIONS_SMA_DIR, DEFINITIONS_DIR, GENERATED_DIR, TAGGED_DIR):
                target_bytes = (d / fname).read_bytes()
                self.assertEqual(
                    prod_bytes,
                    target_bytes,
                    f"File {fname} in {d.name} is not byte-identical to production/{fname}",
                )

    def test_five_apollo_pillars_compliance(self) -> None:
        """Verify each SMA file strictly complies with all 5 Apollo Pillars."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            is_valid, issues = verify_file_pillars(fpath)
            self.assertTrue(
                is_valid,
                f"File {fname} violated 5 Pillars:\n" + "\n".join(issues),
            )

    def test_protocol_header_and_structure(self) -> None:
        """Verify protocol header begins with Modbus; and has 11 columns with canonical manufacturer SMA."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                reader = list(csv.reader(fp, delimiter=";"))

            self.assertGreaterEqual(len(reader), 10, f"File {fname} has too few rows")
            header = reader[0]
            self.assertEqual(len(header), 11, f"Header of {fname} must have 11 columns")
            self.assertEqual(header[0], "Modbus", f"Header protocol of {fname} must be 'Modbus'")
            self.assertEqual(header[2], "SMA", f"Header manufacturer of {fname} must be 'SMA'")

    def test_sequential_indexing_and_row_integrity(self) -> None:
        """Verify sequential index 1..N, valid FCs, decimal addresses, and valid action codes."""
        valid_fcs = {"1", "2", "3", "4"}
        valid_actions = {"0", "1", "4", "9", "10"}

        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags_in_file: set[str] = set()
            for idx, r in enumerate(rows, start=1):
                self.assertEqual(len(r), 11, f"Row {idx} in {fname} must have 11 columns")
                row_idx, fc, addr, dtype, _, name, tag, coef_a, coef_b, unit, action = r

                # 1..N index
                self.assertEqual(row_idx, str(idx), f"Index mismatch at row {idx} in {fname}")

                # Function code
                self.assertIn(fc, valid_fcs, f"Invalid FC '{fc}' at row {idx} in {fname}")

                # Action code
                self.assertIn(
                    action, valid_actions, f"Invalid action '{action}' at row {idx} in {fname}"
                )

                # Coefficients must have 6 decimals
                self.assertRegex(
                    coef_a,
                    r"^-?\d+\.\d{6}$",
                    f"CoefA '{coef_a}' at row {idx} in {fname} must have 6 decimals",
                )
                self.assertRegex(
                    coef_b,
                    r"^-?\d+\.\d{6}$",
                    f"CoefB '{coef_b}' at row {idx} in {fname} must have 6 decimals",
                )

                # Name not empty
                self.assertTrue(name.strip(), f"Empty name at row {idx} in {fname}")

                # Canonical Tag uniqueness
                if tag.strip():
                    self.assertNotIn(
                        tag, tags_in_file, f"Duplicate tag '{tag}' at row {idx} in {fname}"
                    )
                    tags_in_file.add(tag)

    def test_domain_telemetry_contracts(self) -> None:
        """Verify essential domain telemetry tags are present based on equipment role."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags = {r[6].strip() for r in rows if r[6].strip()}

            # Inverters and loggers must measure active power and energy yields
            if fname.startswith("INV_") or fname.startswith("COM_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_DAY" in tags, f"Missing energy yield in {fname}"
                )

            # Core 1 and Core 2 inverters must have multi-MPPT monitoring
            if "Core" in fname or "CORE" in fname or "110-60" in fname or "125-70" in fname:
                self.assertIn("I_DC_1", tags, f"Missing I_DC_1 in {fname}")
                self.assertIn("I_DC_6", tags, f"Missing I_DC_6 in {fname}")

            # Battery storage systems must report state of charge and voltage
            if any(k in fname for k in ("Storage", "SmartEnergy", "Island", "SE40")):
                self.assertIn("SOC", tags, f"Missing SOC in storage {fname}")
                self.assertIn("U_BAT", tags, f"Missing U_BAT in storage {fname}")

            # Meters must measure active power, energy, voltage, and current
            if fname.startswith("MET_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in meter {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_IMP" in tags or "E_EXP" in tags,
                    f"Missing energy counter in meter {fname}",
                )
                self.assertIn("U_AC_1", tags, f"Missing U_AC_1 in meter {fname}")
                self.assertIn("I_AC_1", tags, f"Missing I_AC_1 in meter {fname}")

    def test_sma_threephase_telemetry_correctness(self) -> None:
        """Verify that SMA 3-phase telemetry maps registers accurately:
        - 30775 is active power total (P_AC), NOT rating 30231
        - 30231 is rated/limit capacity (P_RATED or empty), never P_AC
        - 30807 is reactive power phase L1 (Q_AC_1), never S_AC
        - 30813 is apparent power total (S_AC), never PF
        - Power factor (PF) is mapped to cos phi registers (30821, 30825, 30949, 31221, or 40092)
        """
        stpx_path = PRODUCTION_DIR / "INV_SMA_STP_X.csv"
        with open(stpx_path, encoding="utf-8", newline="") as fp:
            rows = {r[2].strip(): (r[6].strip(), r[5].strip(), r[9].strip()) for r in csv.reader(fp, delimiter=";")}

        self.assertEqual(rows.get("30775", ("", "", ""))[0], "P_AC")
        self.assertNotEqual(rows.get("30231", ("", "", ""))[0], "P_AC")
        self.assertEqual(rows.get("30805", ("", "", ""))[0], "Q_AC")
        self.assertEqual(rows.get("30807", ("", "", ""))[0], "Q_AC_1")
        self.assertEqual(rows.get("30809", ("", "", ""))[0], "Q_AC_2")
        self.assertEqual(rows.get("30811", ("", "", ""))[0], "Q_AC_3")
        self.assertEqual(rows.get("30813", ("", "", ""))[0], "S_AC")
        self.assertEqual(rows.get("30815", ("", "", ""))[0], "S_AC_1")
        self.assertEqual(rows.get("30817", ("", "", ""))[0], "S_AC_2")
        self.assertEqual(rows.get("30819", ("", "", ""))[0], "S_AC_3")
        self.assertEqual(rows.get("30949", ("", "", ""))[0], "PF")

        # Check across all 3-phase SMA inverter models
        for fname in (
            "INV_SMA_STP_X.csv",
            "INV_SMA_SunnyTripower_X.csv",
            "INV_SMA_Generic.csv",
            "INV_SMA_SC_Webbox.csv",
            "INV_SMA_STP_5-10-SE40.csv",
            "INV_SMA_SunnyTripower_SmartEnergy.csv",
        ):
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                addr_map = {r[2].strip(): r[6].strip() for r in csv.reader(fp, delimiter=";")}

            self.assertEqual(addr_map.get("30775"), "P_AC", f"30775 in {fname} must be P_AC")
            self.assertNotEqual(addr_map.get("30231"), "P_AC", f"30231 in {fname} must not be P_AC")
            if "30807" in addr_map:
                self.assertEqual(addr_map.get("30807"), "Q_AC_1", f"30807 in {fname} must be Q_AC_1")
            if "30813" in addr_map:
                self.assertEqual(addr_map.get("30813"), "S_AC", f"30813 in {fname} must be S_AC")

    def test_sma_register_addressing_conventions(self) -> None:
        """Verify SMA address space conventions:
        1. SMA Standard Telemetry: base-30000 offset (30001..39999).
        2. SMA Controls & SunSpec Models: base-40000 offset (40001..49999).
        3. Classic Sunny Central: 0-based relative holding offsets (0..500).
        4. Valid Webdyn data types and integer/length address formatting.
        """
        valid_types = {
            "U16",
            "I16",
            "S16",
            "U32",
            "I32",
            "S32",
            "U64",
            "I64",
            "F32",
            "FLOAT32",
            "STRING",
        }

        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            for idx, r in enumerate(rows, start=1):
                addr = r[2].strip()
                dtype = r[3].strip()
                tag = r[6].strip()
                action = r[10].strip()

                self.assertIn(
                    dtype, valid_types, f"Invalid data type '{dtype}' in {fname} row {idx}"
                )
                if tag.startswith("CMD_"):
                    self.assertIn(
                        action,
                        ("0", "1"),
                        f"Command tag '{tag}' must have action 0 or 1 in {fname} row {idx}",
                    )

                # Address format check: either integer or integer_length for STRING
                addr_num = addr.split("_")[0]
                self.assertTrue(
                    addr_num.isdigit(),
                    f"Address '{addr}' in {fname} row {idx} is not numeric",
                )

                val_addr = int(addr_num)
                # Verify address belongs to one of the 3 recognized SMA address schemes:
                # - Classic Sunny Central 0-based: 0 <= addr <= 500
                # - SMA Native telemetry / parameters: 30000 <= addr < 40000
                # - SMA Controls / SunSpec Interoperability: 40000 <= addr < 50000
                is_sc_classic = any(k in fname for k in ("SC2xxx", "SunnyCentral", "SUNNY_CENTRAL"))
                valid_range = (
                    (is_sc_classic and 0 <= val_addr <= 500)
                    or (30000 <= val_addr < 40000)
                    or (40000 <= val_addr < 50000)
                )
                self.assertTrue(
                    valid_range,
                    f"Address {addr} in {fname} does not conform to SMA addressing conventions",
                )

    def test_sma_scale_factors_and_clean_units(self) -> None:
        """Verify canonical SMA scale factors (temperatures -1, voltages/frequency -2, currents -3) and pure units."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            for idx, r in enumerate(rows, start=1):
                addr = r[2].strip()
                tag = r[6].strip()
                ca = float(r[7].strip())
                unit = r[9].strip()

                # Character purity: no duplicate degree signs, no mojibake
                self.assertNotIn("°°", unit, f"Corrupted unit '{unit}' at row {idx} in {fname}")
                self.assertNotIn("-¦", unit, f"Corrupted unit '{unit}' at row {idx} in {fname}")
                self.assertNotIn(
                    "MA°", unit, f"Corrupted MAC unit '{unit}' at row {idx} in {fname}"
                )
                self.assertNotIn(
                    "PP°", unit, f"Corrupted PPC unit '{unit}' at row {idx} in {fname}"
                )

                # Scale factor conventions for SMA native registers (30000 series)
                if addr.isdigit() and 30000 <= int(addr) < 40000:
                    if tag.startswith("TEMP") and unit == "°C":
                        self.assertIn(
                            ca,
                            (0.1, 1.0),
                            f"Unexpected temperature scale factor {ca} in {fname} at {addr}",
                        )
                    elif tag.startswith("U_AC_") and unit == "V":
                        self.assertIn(
                            ca,
                            (0.01, 0.1, 1.0),
                            f"Unexpected AC voltage scale factor {ca} in {fname} at {addr}",
                        )
                    elif tag == "FREQ" and unit in ("Hz", "HZ"):
                        self.assertIn(
                            ca,
                            (0.01, 1.0),
                            f"Unexpected frequency scale factor {ca} in {fname} at {addr}",
                        )
                    elif tag.startswith("I_AC_") and unit == "A":
                        self.assertIn(
                            ca,
                            (0.001, 0.01, 1.0),
                            f"Unexpected AC current scale factor {ca} in {fname} at {addr}",
                        )
                    elif tag == "PF":
                        self.assertIn(
                            ca,
                            (0.001, 0.01, 1.0),
                            f"Unexpected power factor scale factor {ca} in {fname} at {addr}",
                        )

    def test_sma_generator_strict_multi_run_idempotency(self) -> None:
        """Verify that DefFileGenerator.sma_definitions_engine.generate_all_sma_definitions is 100% idempotent."""
        from DefFileGenerator.sma_definitions_engine import generate_all_sma_definitions

        # Initial snapshot
        first_run = generate_all_sma_definitions(REPO_ROOT)
        first_bytes = {f: f.read_bytes() for f in first_run.values()}

        # Second run
        second_run = generate_all_sma_definitions(REPO_ROOT)
        for fname, fpath in second_run.items():
            self.assertEqual(
                first_bytes[fpath],
                fpath.read_bytes(),
                f"Multi-run idempotency violated for {fname}",
            )


    def test_sma_stp_125_70_core2_specifications(self) -> None:
        """Verify SMA Sunny Tripower CORE2 125kW (STP 125-70) specific domain contracts."""
        fname = "INV_SMA_STP_125-70.csv"
        fpath = PRODUCTION_DIR / fname
        with open(fpath, encoding="utf-8", newline="") as fp:
            reader = list(csv.reader(fp, delimiter=";"))

        # Header check
        header = reader[0]
        self.assertEqual(header[0], "Modbus")
        self.assertEqual(header[1], "Inverter")
        self.assertEqual(header[2], "SMA")
        self.assertEqual(header[3], "STP_125-70")

        # Exact register count: 464 data rows + 1 header = 465 lines
        self.assertEqual(len(reader), 465)

        # Extract address map and tag map
        addr_to_row = {r[2].strip(): r for r in reader[1:]}
        tags = {r[6].strip() for r in reader[1:] if r[6].strip()}

        # Canonical single tags
        expected_core_tags = [
            "P_AC",
            "P_AC_1",
            "P_AC_2",
            "P_AC_3",
            "Q_AC",
            "Q_AC_1",
            "Q_AC_2",
            "Q_AC_3",
            "S_AC",
            "S_AC_1",
            "S_AC_2",
            "S_AC_3",
            "FREQ",
            "PF",
            "E_TOT",
            "E_DAY",
            "U_AC_1",
            "U_AC_2",
            "U_AC_3",
            "U_AC_L1L2",
            "U_AC_L2L3",
            "U_AC_L3L1",
            "I_AC",
            "I_AC_1",
            "I_AC_2",
            "I_AC_3",
            "U_DC",
            "P_DC",
            "I_DC",
            "TEMP",
            "STATE",
            "MODEL",
            "SERIAL",
            "MANUFACTURER",
            "CMD_P_ENA",
            "CMD_P_AC_SET_PERCENT",
            "CMD_P_AC_SET_ABS",
        ]
        for t in expected_core_tags:
            self.assertIn(t, tags, f"Missing canonical tag {t} in {fname}")

        # Multi-MPPT tracking across 11 trackers
        for i in range(1, 12):
            self.assertIn(f"U_DC_{i}", tags, f"Missing MPPT voltage U_DC_{i}")
            self.assertIn(f"I_DC_{i}", tags, f"Missing MPPT current I_DC_{i}")
            self.assertIn(f"P_DC_{i}", tags, f"Missing MPPT power P_DC_{i}")

        # String currents 1..12
        for s in range(1, 13):
            self.assertIn(f"I_DC_STRING_{s}", tags, f"Missing string current I_DC_STRING_{s}")

        # Key registers verification
        self.assertEqual(addr_to_row["30001"][5], "SMA Modbus profile version")
        self.assertEqual(addr_to_row["30003"][5], "SUSyID module")
        self.assertEqual(addr_to_row["30005"][5], "Serial number")
        self.assertEqual(addr_to_row["30193"][5], "System time")
        self.assertEqual(addr_to_row["30053"][6], "MODEL")
        self.assertEqual(addr_to_row["30055"][6], "MANUFACTURER")
        self.assertEqual(addr_to_row["30057"][6], "SERIAL")
        self.assertEqual(addr_to_row["30217"][6], "STATE_1")
        self.assertEqual(addr_to_row["30247"][6], "ALARM_1")
        self.assertEqual(addr_to_row["30513"][6], "E_TOT")
        self.assertEqual(addr_to_row["30775"][6], "P_AC")
        self.assertEqual(addr_to_row["30803"][6], "FREQ")
        self.assertEqual(addr_to_row["31497"][6], "Q_AC")
        self.assertEqual(addr_to_row["30813"][6], "S_AC")
        self.assertEqual(addr_to_row["34113"][6], "TEMP")
        self.assertEqual(addr_to_row["34113"][7], "0.100000")
        self.assertEqual(addr_to_row["34113"][9], "°C")


if __name__ == "__main__":
    unittest.main()
