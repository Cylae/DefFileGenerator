"""
Unit and integration tests for Dataset Learning and Definition Audit Engine.
Validates empirical knowledge mining from definition datasets, structural Modbus audits,
semantic SunPM tag validation, automated repairing, and CLI subcommands.
"""

from __future__ import annotations

import csv
import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from DefFileGenerator.knowledge_miner import (
    DatasetKnowledge,
    audit_definition_file,
    audit_definitions_directory,
    auto_fix_definition_file,
    clean_variable_key,
    infer_sunpm_tag,
    train_knowledge_from_definitions,
)
from DefFileGenerator.main import _run_cli


class TestDatasetLearningAndAudit(unittest.TestCase):
    """Test suite for dataset learning, definition auditing, and auto-repairing."""

    def test_clean_variable_key(self) -> None:
        """Verifies token normalization strips prefixes and punctuation."""
        self.assertEqual(clean_variable_key("Inverter Active Power"), "active power")
        self.assertEqual(clean_variable_key("INV_Grid_Frequency (Hz)"), "grid frequency")
        self.assertEqual(clean_variable_key("MPPT2 Current [A]"), "mppt2 current")
        self.assertEqual(clean_variable_key(""), "")
        self.assertEqual(clean_variable_key("Daily Yield"), "daily yield")

    def test_dataset_knowledge_record_and_query(self) -> None:
        """Verifies learning records and lookup in DatasetKnowledge."""
        knowledge = DatasetKnowledge()
        knowledge.record_variable(
            name="Inverter Active Power Output",
            tag="P_AC",
            dtype="U32",
            unit="W",
            action="4",
        )
        knowledge.record_variable(
            name="Inverter Active Power Output",
            tag="P_AC",
            dtype="U32",
            unit="kW",
            action="4",
        )

        learned_tag = knowledge.get_learned_tag("Inverter Active Power Output")
        self.assertEqual(learned_tag, "P_AC")

        # Key normalization lookup
        self.assertEqual(knowledge.get_learned_tag("Active Power Output"), "P_AC")

        # Units and types
        units = knowledge.get_metric_units("active power output")
        self.assertIn("W", units)
        self.assertIn("kW", units)

        dtypes = knowledge.get_metric_types("active power output")
        self.assertIn("U32", dtypes)

        actions = knowledge.get_metric_actions("active power output")
        self.assertIn("4", actions)

    def test_dataset_knowledge_serialization(self) -> None:
        """Verifies serialization to and from dictionary."""
        k1 = DatasetKnowledge()
        k1.record_variable("Grid Voltage L1", "U_AC_1", "U16", "V", "4")
        k1.total_definitions_mined = 5
        k1.total_registers_mined = 120

        data = k1.to_dict()
        k2 = DatasetKnowledge.from_dict(data)

        self.assertEqual(k2.total_definitions_mined, 5)
        self.assertEqual(k2.total_registers_mined, 120)
        self.assertEqual(k2.get_learned_tag("Grid Voltage L1"), "U_AC_1")
        self.assertEqual(k2.tag_counts.get("U_AC_1"), 1)

    def test_train_knowledge_from_definitions(self) -> None:
        """Verifies training against repository Definitions/ directory."""
        definitions_dir = Path("Definitions")
        if not definitions_dir.is_dir():
            definitions_dir = Path(__file__).parent.parent.parent / "Definitions"

        self.assertTrue(definitions_dir.is_dir(), "Definitions directory must exist")

        knowledge = train_knowledge_from_definitions(
            directories=definitions_dir, force_retrain=True
        )
        self.assertGreaterEqual(knowledge.total_definitions_mined, 120)
        self.assertGreaterEqual(knowledge.total_registers_mined, 15000)
        self.assertGreaterEqual(len(knowledge.name_to_tag), 500)

        # Confirm empirical lookups
        self.assertIn("P_AC", knowledge.tag_counts)
        self.assertIn("E_TOT", knowledge.tag_counts)

    def test_infer_sunpm_tag_with_dataset_knowledge(self) -> None:
        """Verifies infer_sunpm_tag uses learned knowledge before falling back to rules."""
        custom_knowledge = DatasetKnowledge()
        custom_knowledge.record_variable("SuperSpecialEmpiricalMetric", "CUSTOM_TAG_TEST")

        inferred = infer_sunpm_tag(
            variable_name="SuperSpecialEmpiricalMetric",
            knowledge=custom_knowledge,
        )
        self.assertEqual(inferred, "CUSTOM_TAG_TEST")

    def test_audit_definition_file_valid(self) -> None:
        """Verifies that an audited valid file produces no errors."""
        definitions_dir = Path("Definitions")
        if not definitions_dir.is_dir():
            definitions_dir = Path(__file__).parent.parent.parent / "Definitions"

        target_file = definitions_dir / "INV_CyberPower_GENERIC.csv"
        self.assertTrue(target_file.is_file())

        report = audit_definition_file(target_file, strict_overlap=True)
        self.assertTrue(report.is_valid)
        self.assertEqual(len(report.errors), 0)
        self.assertGreater(report.register_count, 100)

    def test_audit_definition_file_detects_defects(self) -> None:
        """Verifies audit engine flags overlaps, legacy tags, incoherent units, and invalid actions."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "defective_definition.csv"
            rows = [
                ["ID", "Inverter", "DefectiveModel"],
                ["0", "4", "100", "U32", "0", "Total Power", "RealPower", "1", "0", "V", "4"],
                ["1", "4", "101", "U16", "0", "Overlap Register", "", "1", "0", "kW", "4"],
                ["2", "4", "200", "U16", "0", "Inverter Turn On", "CMD_ON", "1", "0", "", "4"],
                ["3", "4", "300", "U16", "0", "Total Energy Yield", "E_TOT", "1", "0", "kW", "4"],
            ]
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(rows)

            report = audit_definition_file(csv_path, strict_overlap=True, strict_actions=True)
            self.assertFalse(report.is_valid)

            error_codes = [issue.code for issue in report.errors]
            warning_codes = [issue.code for issue in report.warnings]

            # Overlap at 101 (since U32 at 100 occupies 100 and 101)
            self.assertIn("ADDRESS_OVERLAP", error_codes)

            # Strict action check: CMD_ON with action 4
            self.assertIn("INCOHERENT_ACTION", error_codes)

            # Legacy tag 'RealPower' -> warning
            self.assertIn("LEGACY_TAG", warning_codes)

            # Incoherent unit: E_TOT with 'kW' instead of energy unit -> warning
            self.assertIn("INCOHERENT_UNIT", warning_codes)

    def test_auto_fix_definition_file(self) -> None:
        """Verifies auto_fix sanitizes legacy tags, infers known tags, and removes duplicate rows."""
        with tempfile.TemporaryDirectory() as tmpdir:
            in_csv = Path(tmpdir) / "input.csv"
            out_csv = Path(tmpdir) / "sanitized.csv"

            rows = [
                ["ID", "Inverter", "AutoFixModel"],
                ["0", "4", "100", "U16", "0", "Grid Voltage L1", "RealPower", "1", "0", "V", "4"],
                ["1", "4", "100", "U16", "0", "Grid Voltage L1", "RealPower", "1", "0", "V", "4"],
                ["2", "4", "200", "U16", "0", "Inverter Grid Frequency", "", "1", "0", "Hz", "4"],
            ]
            with open(in_csv, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(rows)

            counts = auto_fix_definition_file(in_csv, out_csv)
            self.assertEqual(counts["duplicates_removed"], 1)
            self.assertEqual(counts["tags_normalized"], 1)
            self.assertEqual(counts["tags_inferred"], 1)

            # Verify sanitized output content
            with open(out_csv, newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                sanitized_rows = list(reader)

            # Header + 2 rows (duplicate eliminated)
            self.assertEqual(len(sanitized_rows), 3)
            # RealPower normalized to P_AC
            self.assertEqual(sanitized_rows[1][6], "P_AC")
            # Inverter Grid Frequency inferred as FREQ
            self.assertEqual(sanitized_rows[2][6], "FREQ")

    def test_audit_definitions_directory(self) -> None:
        """Verifies directory auditing aggregates multiple definition files."""
        definitions_dir = Path("Definitions")
        if not definitions_dir.is_dir():
            definitions_dir = Path(__file__).parent.parent.parent / "Definitions"

        reports = audit_definitions_directory(definitions_dir, strict_overlap=False)
        self.assertGreaterEqual(len(reports), 120)
        total_regs = sum(r.register_count for r in reports.values())
        self.assertGreaterEqual(total_regs, 15000)

    def test_cli_train_command(self) -> None:
        """Verifies `deffilegen train --summary --export-json` CLI subcommand."""
        with tempfile.TemporaryDirectory() as tmpdir:
            json_export = Path(tmpdir) / "learned.json"
            _run_cli(["train", "--summary", "--export-json", str(json_export)])

            self.assertTrue(json_export.is_file())
            with open(json_export, encoding="utf-8") as f:
                data = json.load(f)
            self.assertIn("name_to_tag", data)
            self.assertIn("tag_counts", data)
            self.assertGreater(data["total_registers_mined"], 10000)

    def test_cli_audit_command_json(self) -> None:
        """Verifies `deffilegen audit <file> --json` CLI output format."""
        definitions_dir = Path("Definitions")
        if not definitions_dir.is_dir():
            definitions_dir = Path(__file__).parent.parent.parent / "Definitions"

        target_file = definitions_dir / "INV_CyberPower_GENERIC.csv"

        buf = io.StringIO()
        with redirect_stdout(buf):
            _run_cli(["audit", str(target_file), "--json"])

        output_str = buf.getvalue()
        self.assertTrue(output_str.strip().startswith("{"))
        data = json.loads(output_str)
        self.assertTrue(data["is_valid"])
        self.assertEqual(data["stats"]["errors"], 0)


if __name__ == "__main__":
    unittest.main()
