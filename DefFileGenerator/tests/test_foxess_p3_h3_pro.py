#!/usr/bin/env python3
"""
Unit and integration tests for Fox ESS P3 Pro and H3 Pro definitions:
- Specific P3 Pro definitions (Inverter, Meter, Battery) + P3_Pro/README.txt
- Specific H3 Pro definitions (Inverter, Meter, Battery) + H3_Pro/README.txt
- Universal combined series (FOXESS_P3_H3_PRO_INVERTER, METER, BATTERY, FoxESS_P3_H3_PRO)
- Local delivery directory FoxESS_P3_H3_Pro_Definitions/
- Technical audit reports (Markdown & JSON)
"""

import csv
import json
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
LOCAL_DIR = REPO_ROOT / "FoxESS_P3_H3_Pro_Definitions"
P3_PRO_DIR = LOCAL_DIR / "P3_Pro"
H3_PRO_DIR = LOCAL_DIR / "H3_Pro"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"
GEN_DIR = REPO_ROOT / "generated_definitions"
DOCS_DIR = REPO_ROOT / "docs"


class TestFoxESSP3H3Pro(unittest.TestCase):
    """Test suite verifying 100% compliance of Fox ESS P3 Pro & H3 Pro definitions."""

    generator: Generator

    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = Generator()

    def test_p3_pro_dedicated_folder_and_readme(self) -> None:
        """Verify FoxESS P3 Pro dedicated folder, files, and README."""
        self.assertTrue(P3_PRO_DIR.exists(), f"Missing {P3_PRO_DIR}")
        readme_path = P3_PRO_DIR / "README.txt"
        self.assertTrue(readme_path.exists(), "Missing P3_Pro/README.txt")
        readme_content = readme_path.read_text(encoding="utf-8")
        self.assertIn("FOXESS GAMME P3 PRO", readme_content)
        self.assertIn("FOXESS_P3_PRO_INVERTER.csv", readme_content)
        self.assertIn("FOXESS_P3_PRO_METER.csv", readme_content)
        self.assertIn("FOXESS_P3_PRO_BATTERY.csv", readme_content)

        inv_file = P3_PRO_DIR / "FOXESS_P3_PRO_INVERTER.csv"
        met_file = P3_PRO_DIR / "FOXESS_P3_PRO_METER.csv"
        bat_file = P3_PRO_DIR / "FOXESS_P3_PRO_BATTERY.csv"

        for f, model, eq in [
            (inv_file, "P3 Pro", "Inverter"),
            (met_file, "P3 Pro", "Meter"),
            (bat_file, "P3 Pro", "Battery"),
        ]:
            self.assertTrue(f.exists(), f"Missing {f}")
            rep = self.generator.validate_csv_detailed(str(f), strict=False, strict_overlap=False)
            self.assertTrue(
                rep.is_valid,
                f"{f.name} validation failed: {[i.message for i in rep.issues if i.severity == 'ERROR']}",
            )
            with open(f, encoding="utf-8-sig") as fp:
                hdr = next(csv.reader(fp, delimiter=";"))
                self.assertEqual(hdr[1], eq)
                self.assertEqual(hdr[2], "FoxESS")
                self.assertEqual(hdr[3], model)

    def test_h3_pro_dedicated_folder_and_readme(self) -> None:
        """Verify FoxESS H3 Pro dedicated folder, files, and README."""
        self.assertTrue(H3_PRO_DIR.exists(), f"Missing {H3_PRO_DIR}")
        readme_path = H3_PRO_DIR / "README.txt"
        self.assertTrue(readme_path.exists(), "Missing H3_Pro/README.txt")
        readme_content = readme_path.read_text(encoding="utf-8")
        self.assertIn("FOXESS GAMME H3 PRO", readme_content)
        self.assertIn("FOXESS_H3_PRO_INVERTER.csv", readme_content)
        self.assertIn("FOXESS_H3_PRO_METER.csv", readme_content)
        self.assertIn("FOXESS_H3_PRO_BATTERY.csv", readme_content)

        inv_file = H3_PRO_DIR / "FOXESS_H3_PRO_INVERTER.csv"
        met_file = H3_PRO_DIR / "FOXESS_H3_PRO_METER.csv"
        bat_file = H3_PRO_DIR / "FOXESS_H3_PRO_BATTERY.csv"

        for f, model, eq in [
            (inv_file, "H3 Pro", "Inverter"),
            (met_file, "H3 Pro", "Meter"),
            (bat_file, "H3 Pro", "Battery"),
        ]:
            self.assertTrue(f.exists(), f"Missing {f}")
            rep = self.generator.validate_csv_detailed(str(f), strict=False, strict_overlap=False)
            self.assertTrue(
                rep.is_valid,
                f"{f.name} validation failed: {[i.message for i in rep.issues if i.severity == 'ERROR']}",
            )
            with open(f, encoding="utf-8-sig") as fp:
                hdr = next(csv.reader(fp, delimiter=";"))
                self.assertEqual(hdr[1], eq)
                self.assertEqual(hdr[2], "FoxESS")
                self.assertEqual(hdr[3], model)

    def test_root_readmes_and_reports(self) -> None:
        """Verify root README.txt, README.md, and technical reports exist and are complete."""
        root_txt = LOCAL_DIR / "README.txt"
        root_md = LOCAL_DIR / "README.md"
        doc_report = DOCS_DIR / "RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md"
        tagged_report = TAGGED_DIR / "RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md"
        json_report = TAGGED_DIR / "foxess_p3_h3_pro_report.json"

        self.assertTrue(root_txt.exists(), "Missing root README.txt")
        self.assertTrue(root_md.exists(), "Missing root README.md")
        self.assertTrue(doc_report.exists(), "Missing docs technical report")
        self.assertTrue(tagged_report.exists(), "Missing tagged_definitions technical report")
        self.assertTrue(json_report.exists(), "Missing json metrics report")

        txt_content = root_txt.read_text(encoding="utf-8")
        self.assertIn("FOXESS GAMMES P3 PRO & H3 PRO", txt_content)
        self.assertIn("P3_Pro/", txt_content)
        self.assertIn("H3_Pro/", txt_content)

        data = json.loads(json_report.read_text(encoding="utf-8"))
        self.assertTrue(data["summary"]["all_valid"])
        self.assertEqual(data["manufacturer"], "Fox ESS")

    def test_inverter_canonical_telemetry_tags(self) -> None:
        """Verify Inverter canonical tags: P_DC, P_AC, Q_AC, FREQ, PF, P_RATED, CMD_ON, CMD_OFF."""
        inv_file = LOCAL_DIR / "FOXESS_P3_H3_PRO_INVERTER.csv"
        self.assertTrue(inv_file.exists(), f"Missing {inv_file}")

        with open(inv_file, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter=";"))[1:]

        addr_to_tag = {r[2].split("_")[0]: r[6].strip() for r in rows if r[6].strip()}
        exact_addr_to_tag = {r[2].strip(): r[6].strip() for r in rows if r[6].strip()}

        self.assertEqual(
            addr_to_tag.get("39118"), "P_DC", "39118 must be tagged P_DC (Total PV input power)"
        )
        self.assertEqual(
            addr_to_tag.get("39134"), "P_AC", "39134 must be tagged P_AC (Active power)"
        )
        self.assertEqual(
            addr_to_tag.get("39136"), "Q_AC", "39136 must be tagged Q_AC (Reactive power)"
        )
        self.assertEqual(addr_to_tag.get("39138"), "PF", "39138 must be tagged PF (Power Factor)")
        self.assertEqual(
            addr_to_tag.get("39139"), "FREQ", "39139 must be tagged FREQ (Grid Frequency)"
        )
        self.assertEqual(addr_to_tag.get("39053"), "P_RATED", "39053 must be tagged P_RATED")
        self.assertEqual(addr_to_tag.get("30000"), "MODEL", "30000 must be tagged MODEL")
        self.assertEqual(addr_to_tag.get("30016"), "SERIAL", "30016 must be tagged SERIAL")
        self.assertEqual(
            addr_to_tag.get("30032"), "MANUFACTURER", "30032 must be tagged MANUFACTURER"
        )
        self.assertEqual(exact_addr_to_tag.get("49077"), "CMD_ON", "49077 must be CMD_ON")
        self.assertEqual(exact_addr_to_tag.get("49078"), "CMD_OFF", "49078 must be CMD_OFF")

        # Verify alarm bits do NOT have analog telemetry tags
        for r in rows:
            addr = r[2]
            if addr.startswith("39067_") or addr.startswith("39068_") or addr.startswith("39069_"):
                self.assertNotIn(
                    r[6],
                    ("FREQ", "I_DC", "R_ISO", "P_AC", "Q_AC", "P_DC"),
                    f"Alarm register {addr} must not steal analog telemetry tags",
                )
                self.assertEqual(r[10], "9", f"Alarm register {addr} must have Action = 9")

    def test_meter_canonical_telemetry_tags(self) -> None:
        """Verify Meter definition 3-phase grid telemetry and mandatory tags."""
        met_file = LOCAL_DIR / "FOXESS_P3_H3_PRO_METER.csv"
        self.assertTrue(met_file.exists(), f"Missing {met_file}")

        with open(met_file, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter=";"))[1:]

        tags = {r[6].strip() for r in rows if r[6].strip()}
        mandatory_tags = {
            "SERIAL",
            "MANUFACTURER",
            "MODEL",
            "U_AC_1",
            "U_AC_2",
            "U_AC_3",
            "I_AC_1",
            "I_AC_2",
            "I_AC_3",
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
            "PF",
            "FREQ",
        }
        for mt in mandatory_tags:
            self.assertIn(mt, tags, f"Meter must contain tag '{mt}'")

    def test_battery_canonical_telemetry_tags(self) -> None:
        """Verify Battery definition BMS telemetry, unique tags, and commands."""
        bat_file = LOCAL_DIR / "FOXESS_P3_H3_PRO_BATTERY.csv"
        self.assertTrue(bat_file.exists(), f"Missing {bat_file}")

        with open(bat_file, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter=";"))[1:]

        tags = {r[6].strip() for r in rows if r[6].strip()}
        mandatory_tags = {
            "MODEL",
            "SERIAL",
            "MANUFACTURER",
            "U_DC",
            "I_DC",
            "TEMP",
            "SOC",
            "SOH",
            "TEMP_MAX",
            "TEMP_MIN",
            "U_CELL_MAX",
            "U_CELL_MIN",
            "ALARM_1",
        }
        for mt in mandatory_tags:
            self.assertIn(mt, tags, f"Battery must contain tag '{mt}'")

        # Verify uniqueness of critical tags
        tag_counts: dict[str, int] = {}
        for r in rows:
            t = r[6].strip()
            if t:
                tag_counts[t] = tag_counts.get(t, 0) + 1
        for unique_tag in ("MODEL", "SERIAL", "MANUFACTURER", "SOC", "SOH", "TEMP", "U_DC", "I_DC"):
            if unique_tag in tag_counts:
                self.assertEqual(
                    tag_counts[unique_tag],
                    1,
                    f"Tag '{unique_tag}' must be unique, found {tag_counts[unique_tag]} occurrences",
                )
                self.assertEqual(
                    tag_counts[unique_tag],
                    1,
                    f"Tag '{unique_tag}' must be unique, found {tag_counts[unique_tag]} occurrences",
                )

    def test_arbitration_register_parity_p3_vs_h3(self) -> None:
        """Factual arbitration test: verify that P3 Pro and H3 Pro share identical registers."""
        for eq_type, p3_name, h3_name in [
            ("Inverter", "FOXESS_P3_PRO_INVERTER.csv", "FOXESS_H3_PRO_INVERTER.csv"),
            ("Meter", "FOXESS_P3_PRO_METER.csv", "FOXESS_H3_PRO_METER.csv"),
            ("Battery", "FOXESS_P3_PRO_BATTERY.csv", "FOXESS_H3_PRO_BATTERY.csv"),
        ]:
            p3_file = P3_PRO_DIR / p3_name
            h3_file = H3_PRO_DIR / h3_name
            self.assertTrue(p3_file.exists(), f"Missing {p3_file}")
            self.assertTrue(h3_file.exists(), f"Missing {h3_file}")

            with (
                open(p3_file, encoding="utf-8-sig") as f1,
                open(h3_file, encoding="utf-8-sig") as f2,
            ):
                r1 = list(csv.reader(f1, delimiter=";"))
                r2 = list(csv.reader(f2, delimiter=";"))

            # Headers differ only in Model field (column 3: 'P3 Pro' vs 'H3 Pro')
            self.assertEqual(r1[0][1], r2[0][1], f"DeviceType mismatch for {eq_type}")
            self.assertEqual(r1[0][2], r2[0][2], f"Manufacturer mismatch for {eq_type}")
            self.assertEqual(r1[0][3], "P3 Pro")
            self.assertEqual(r2[0][3], "H3 Pro")

            # Data rows must be 100% identical in address, FC, type, tag, Ca, Cb, unit, action
            self.assertEqual(len(r1), len(r2), f"Row count mismatch between P3 and H3 {eq_type}")
            for idx in range(1, len(r1)):
                row_p3 = r1[idx]
                row_h3 = r2[idx]
                # Compare FC, Address, DataType, Tag, Ca, Cb, Unit, Action
                self.assertEqual(row_p3[1], row_h3[1], f"FC mismatch at row {idx}")
                self.assertEqual(row_p3[2], row_h3[2], f"Address mismatch at row {idx}")
                self.assertEqual(row_p3[3], row_h3[3], f"DataType mismatch at row {idx}")
                self.assertEqual(row_p3[6], row_h3[6], f"Tag mismatch at row {idx}")
                self.assertEqual(row_p3[7], row_h3[7], f"Ca mismatch at row {idx}")
                self.assertEqual(row_p3[8], row_h3[8], f"Cb mismatch at row {idx}")
                self.assertEqual(row_p3[10], row_h3[10], f"Action mismatch at row {idx}")

    def test_all_in_one_unified_definition(self) -> None:
        """Verify the All-In-One definition combines Inverter, Meter, and Battery with 0 conflict."""
        aio_file = LOCAL_DIR / "FOXESS_P3_H3_PRO_ALL_IN_ONE.csv"
        self.assertTrue(aio_file.exists(), f"Missing {aio_file}")

        rep = self.generator.validate_csv_detailed(
            str(aio_file), strict=False, strict_overlap=False
        )
        self.assertTrue(
            rep.is_valid,
            f"All-In-One validation failed: {[i.message for i in rep.issues if i.severity == 'ERROR']}",
        )
        self.assertEqual(rep.register_count, 1233, "All-In-One must have exactly 1,233 registers")

        with open(aio_file, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            hdr = next(reader)
            rows = list(reader)

        self.assertEqual(hdr[1], "Inverter")
        self.assertEqual(hdr[2], "FoxESS")
        self.assertEqual(hdr[3], "P3 Pro / H3 Pro")

        # Check sequential device IDs (col 0)
        for i, row in enumerate(rows, 1):
            self.assertEqual(row[0], str(i), f"Device ID mismatch at row {i}")
            self.assertEqual(row[4], "", f"ScaleFactor (col 5) must be empty at row {i}")

        # Check key registers from all 3 domains
        addrs = {r[2].split("_")[0] for r in rows}
        self.assertIn("39134", addrs, "Inverter active power must be present")
        self.assertIn("38814", addrs, "Meter active power must be present")
        self.assertIn("37612", addrs, "Battery SoC must be present")

    def test_commun_dedicated_folder_and_readme(self) -> None:
        """Verify the dedicated Commun/ folder and technical arbitration documentation."""
        commun_dir = LOCAL_DIR / "Commun"
        self.assertTrue(commun_dir.exists(), f"Missing {commun_dir}")

        readme_file = commun_dir / "README.txt"
        self.assertTrue(readme_file.exists(), "Missing Commun/README.txt")
        readme_txt = readme_file.read_text(encoding="utf-8")
        self.assertIn("FICHIERS DE DEFINITIONS COMMUNS", readme_txt)
        self.assertIn("FOXESS_P3_H3_PRO_ALL_IN_ONE.csv", readme_txt)

        self.assertIn("FOXESS_P3_H3_PRO_INVERTER.csv", readme_txt)
        self.assertIn("FOXESS_P3_H3_PRO_METER.csv", readme_txt)
        self.assertIn("FOXESS_P3_H3_PRO_BATTERY.csv", readme_txt)

        for fname in [
            "FOXESS_P3_H3_PRO_ALL_IN_ONE.csv",
            "FOXESS_P3_H3_PRO_INVERTER.csv",
            "FOXESS_P3_H3_PRO_METER.csv",
            "FOXESS_P3_H3_PRO_BATTERY.csv",
        ]:
            target = commun_dir / fname
            self.assertTrue(target.exists(), f"Missing {target}")
            rep = self.generator.validate_csv_detailed(
                str(target), strict=False, strict_overlap=False
            )
            self.assertTrue(rep.is_valid, f"{fname} in Commun/ failed validation")

    def test_other_foxess_sources_from_network_share(self) -> None:
        """Verify all other FoxESS equipment sources from network share are properly generated and valid."""
        autres_dir = LOCAL_DIR / "Autres_Gammes_FoxESS"
        self.assertTrue(autres_dir.exists(), f"Missing {autres_dir}")

        readme_file = autres_dir / "README.txt"
        self.assertTrue(readme_file.exists(), "Missing Autres_Gammes_FoxESS/README.txt")

        # Test G-MAX
        gmax_file = autres_dir / "FOXESS_GMAX_STORAGE.csv"
        self.assertTrue(gmax_file.exists(), "Missing FOXESS_GMAX_STORAGE.csv")
        rep_gmax = self.generator.validate_csv_detailed(
            str(gmax_file), strict=False, strict_overlap=False
        )
        self.assertTrue(rep_gmax.is_valid, "GMAX validation failed")
        self.assertGreater(rep_gmax.register_count, 800)

        # Test R-Series Inverter
        r_inv = autres_dir / "FOXESS_R_SERIES_INVERTER.csv"
        self.assertTrue(r_inv.exists(), "Missing FOXESS_R_SERIES_INVERTER.csv")
        rep_rinv = self.generator.validate_csv_detailed(
            str(r_inv), strict=False, strict_overlap=False
        )
        self.assertTrue(rep_rinv.is_valid, "R-Series Inverter validation failed")

        # Test R-Series Battery
        r_bat = autres_dir / "FOXESS_R_SERIES_BATTERY.csv"
        self.assertTrue(r_bat.exists(), "Missing FOXESS_R_SERIES_BATTERY.csv")
        rep_rbat = self.generator.validate_csv_detailed(
            str(r_bat), strict=False, strict_overlap=False
        )
        self.assertTrue(rep_rbat.is_valid, "R-Series Battery validation failed")

        # Test SunSpec RV H3plus
        sun_file = autres_dir / "FOXESS_SUNSPEC_RV_H3PLUS.csv"
        self.assertTrue(sun_file.exists(), "Missing FOXESS_SUNSPEC_RV_H3PLUS.csv")
        rep_sun = self.generator.validate_csv_detailed(
            str(sun_file), strict=False, strict_overlap=False
        )
        self.assertTrue(rep_sun.is_valid, "SunSpec validation failed")


if __name__ == "__main__":
    unittest.main()
