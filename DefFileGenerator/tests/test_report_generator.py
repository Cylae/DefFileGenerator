#!/usr/bin/env python3
"""
Unit and integration tests for DefFileGenerator.report_generator.

Tests cover:
    - Single file analysis and metrics extraction
    - Directory batch analysis and KPI aggregation
    - Markdown report generation (tables, charts, focus sections)
    - JSON metrics export
    - CLI 'report' subcommand dispatch
    - Integration with generate_webdyn_def.py --report flag
"""

import json
import os
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.report_generator import (
    DefinitionFileStats,
    DirectoryReportStats,
    analyze_definition_file,
    analyze_definitions_directory,
    export_report_json,
    generate_directory_technical_report,
    generate_single_definition_report,
)
from generate_webdyn_def import generate_webdyn_definition


class TestReportGenerator(unittest.TestCase):
    """Test suite for the Technical Report Generator module."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

        # Create sample valid definition CSV 1 (Inverter)
        self.csv1 = self.test_dir / "INV_SampleMfg_Inverter100.csv"
        with open(self.csv1, "w", encoding="utf-8-sig", newline="") as f:
            f.write("modbusRTU;Inverter;SampleMfg;Inverter100;;;;;;;\n")
            f.write("1;3;40001;U16;;Active Power;P_AC;1.0;0.0;W;4\n")
            f.write("2;3;40002;U16;;Grid Frequency;FREQ;0.01;0.0;Hz;4\n")
            f.write("3;3;40003;U16;;Cos Phi;PF;0.001;0.0;;4\n")
            f.write("4;3;40004;U16;;Turn On Command;CMD_ON;1.0;0.0;;1\n")
            f.write("5;1;10001;U16;;Remote Stop Coil;CMD_OFF;1.0;0.0;;1\n")

        # Create sample valid definition CSV 2 (Meter)
        self.csv2 = self.test_dir / "MET_SampleMfg_PowerMeter200.csv"
        with open(self.csv2, "w", encoding="utf-8-sig", newline="") as f:
            f.write("modbusRTU;Meter;SampleMfg;PowerMeter200;;;;;;;\n")
            f.write("1;4;30001;U16;;Voltage L1;U_AC_1;0.1;0.0;V;4\n")
            f.write("2;4;30002;U16;;Current L1;I_AC_1;0.01;0.0;A;4\n")
            f.write("3;4;30003;U16;;Total Energy;E_TOT;1.0;0.0;kWh;4\n")

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_analyze_definition_file(self) -> None:
        """Verify file-level metrics extraction and Modbus classification."""
        stats = analyze_definition_file(self.csv1)
        self.assertIsInstance(stats, DefinitionFileStats)
        self.assertEqual(stats.filename, "INV_SampleMfg_Inverter100.csv")
        self.assertEqual(stats.prefix, "INV")
        self.assertEqual(stats.category, "Inverter")
        self.assertEqual(stats.manufacturer, "SampleMfg")
        self.assertEqual(stats.model, "Inverter100")
        self.assertEqual(stats.register_count, 5)
        self.assertEqual(stats.holding_count, 4)
        self.assertEqual(stats.action_1_count, 2)
        self.assertEqual(stats.action_4_count, 3)
        self.assertEqual(stats.action_9_count, 0)
        self.assertEqual(stats.canonical_tag_count, 5)
        self.assertEqual(stats.tag_density, 100.0)
        self.assertTrue(stats.is_valid)

    def test_analyze_definitions_directory(self) -> None:
        """Verify aggregated directory KPIs across multiple definitions."""
        dir_stats = analyze_definitions_directory(self.test_dir)
        self.assertIsInstance(dir_stats, DirectoryReportStats)
        self.assertEqual(dir_stats.total_files, 2)
        self.assertEqual(dir_stats.valid_files, 2)
        self.assertEqual(dir_stats.invalid_files, 0)
        self.assertEqual(dir_stats.total_registers, 8)
        self.assertEqual(dir_stats.total_tags, 8)
        self.assertEqual(dir_stats.holding_registers, 7)
        self.assertEqual(dir_stats.coil_registers, 1)
        self.assertEqual(dir_stats.category_counts.get("INV"), 1)
        self.assertEqual(dir_stats.category_counts.get("MET"), 1)
        self.assertEqual(dir_stats.manufacturer_counts.get("SampleMfg"), 2)

    def test_generate_directory_technical_report(self) -> None:
        """Verify complete Markdown report synthesis and structure."""
        report_path = self.test_dir / "TECHNICAL_REPORT.md"
        content = generate_directory_technical_report(
            self.test_dir, output_report_path=report_path
        )

        self.assertTrue(report_path.is_file())
        self.assertIn("# 📊 RAPPORT TECHNIQUE MAJEUR", content)
        self.assertIn("## 1. Synthèse Exécutive & Indicateurs Clés (KPIs)", content)
        self.assertIn("```mermaid", content)
        self.assertIn("## 2. Analyse Détaillée par Catégorie d'Équipement", content)
        self.assertIn("## 3. Couverture des Principaux Constructeurs Industriels", content)
        self.assertIn("## 4. Cartographie & Distribution des Tags SunPM Canoniques", content)
        self.assertIn("## 5. Focus Systèmes Complexes & Découpage Modulaire", content)
        self.assertIn("## 6. Audit Qualité & Conformité Webdyn Schema", content)
        self.assertIn("INV", content)
        self.assertIn("MET", content)
        self.assertIn("SampleMfg", content)

    def test_generate_single_definition_report(self) -> None:
        """Verify single-file technical report synthesis."""
        out_report = self.test_dir / "INV_single_report.md"
        content = generate_single_definition_report(self.csv1, output_report_path=out_report)

        self.assertTrue(out_report.is_file())
        self.assertIn("RAPPORT TECHNIQUE DE DÉFINITION", content)
        self.assertIn("SampleMfg", content)
        self.assertIn("Inverter100", content)
        self.assertIn("5 registres", content)
        self.assertIn("P_AC", content)

    def test_export_report_json(self) -> None:
        """Verify structured JSON metrics serialization."""
        dir_stats = analyze_definitions_directory(self.test_dir)
        json_path = self.test_dir / "report.json"
        export_report_json(dir_stats, json_path)

        self.assertTrue(json_path.is_file())
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertEqual(data["total_files"], 2)
        self.assertEqual(data["valid_files"], 2)
        self.assertEqual(data["total_registers"], 8)
        self.assertIn("category_counts", data)
        self.assertIn("INV", data["category_counts"])
        self.assertIn("MET", data["category_counts"])

    def test_cli_report_subcommand(self) -> None:
        """Verify command-line execution of 'deffilegen report'."""
        from DefFileGenerator.main import main as cli_main

        out_md = self.test_dir / "cli_report.md"
        out_json = self.test_dir / "cli_report.json"

        # Test directory report via CLI
        cli_main(["report", str(self.test_dir), "-o", str(out_md), "--json", str(out_json)])
        self.assertTrue(out_md.is_file())
        self.assertTrue(out_json.is_file())

        # Test single file report via CLI
        single_md = self.test_dir / "single_file_report.md"
        cli_main(["report", str(self.csv1), "-o", str(single_md)])
        self.assertTrue(single_md.is_file())

    def test_generate_webdyn_def_with_report(self) -> None:
        """Verify generate_webdyn_definition automatically generates report when requested."""
        raw_input = self.test_dir / "registers.csv"
        with open(raw_input, "w", encoding="utf-8") as f:
            f.write("Register,Name,Data Type,Unit,Scale,Access\n")
            f.write("40001,Grid Power,uint16,W,1,R\n")
            f.write("40002,Grid Voltage,uint16,V,0.1,R\n")

        output_csv = self.test_dir / "generated_def.csv"
        report_md = self.test_dir / "gen_report.md"

        success = generate_webdyn_definition(
            input_file=str(raw_input),
            output_file=str(output_csv),
            manufacturer="TestBrand",
            model="ModelX",
            report_file=str(report_md),
        )

        self.assertTrue(success)
        self.assertTrue(output_csv.is_file())
        self.assertTrue(report_md.is_file())

        with open(report_md, "r", encoding="utf-8") as f:
            rep_text = f.read()
        self.assertIn("RAPPORT TECHNIQUE DE DÉFINITION", rep_text)
        self.assertIn("TestBrand", rep_text)


if __name__ == "__main__":
    unittest.main()
