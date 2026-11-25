"""
Apollo-Standard Adversarial Test Suite for Huawei Master Definitions.

Tests full compliance and domain fidelity for:
  - SUN2000-V4 Generic, 2..5KTL, 8..20KTL, 30..40KTL-M3, 33A..42KTL, 50..70KTL-M0, 60KTL-HV, 90..105KTL, V3, All-V3
  - LUNA2000 & LUNA2000 Storage
  - SmartLogger 3000 & SmartLogger V Series
  - Power Meter DTSU666-H
"""

from __future__ import annotations

import csv
from pathlib import Path

from DefFileGenerator.io_utils import find_definition_file

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

HUAWEI_MODELS = [
    ("INV", "SUN2000-V4_Generic"),
    ("INV", "SUN2000-2_5KTL"),
    ("INV", "SUN2000-8_20KTL"),
    ("INV", "SUN2000-30_40KTL-M3"),
    ("INV", "SUN2000-33A_36_42KTL"),
    ("INV", "SUN2000-50_70KTL-M0"),
    ("INV", "SUN2000-60KTL-HV"),
    ("INV", "SUN2000-90_105KTL"),
    ("INV", "SUN2000-V3"),
    ("INV", "SUN2000-All_V3"),
    ("BAT", "LUNA2000"),
    ("BAT", "LUNA2000_Storage"),
    ("COM", "SmartLogger3000"),
    ("COM", "SmartLogger_VSeries"),
    ("MET", "PowerMeter"),
]


class TestHuaweiDefinitionsSuite:
    """Rigorous verification suite for Huawei definition files."""

    def test_all_huawei_definitions_exist(self) -> None:
        """Verify that all target Huawei definition files exist in production/."""
        for pfx, model in HUAWEI_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_Huawei_{model}.csv")
            assert file_path.is_file(), f"Missing definition file for {model} at {file_path}"
            assert file_path.stat().st_size > 500, f"File suspiciously small: {file_path.name}"

    def test_all_huawei_5pillars_strict_certification(self) -> None:
        """Verify that all Huawei definitions satisfy the 5 Webdyn Pillars with strict CRLF."""
        for pfx, model in HUAWEI_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_Huawei_{model}.csv")
            valid, issues = verify_file_pillars(file_path, strict_crlf=True)
            assert valid, f"{file_path.name} failed 5-pillars certification: {issues}"

    def test_sun2000_v4_registers(self) -> None:
        """Verify SUN2000-V4 has authentic registers."""
        file_path = find_definition_file(PROD_DIR, "INV_Huawei_SUN2000-V4_Generic.csv")
        with open(file_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[2] == "Huawei"
            rows = list(reader)
            assert len(rows) > 300

    def test_multi_directory_synchronization(self) -> None:
        """Verify all Huawei definitions are synchronized across target directories."""
        for pfx, model in HUAWEI_MODELS:
            filename = f"{pfx}_Huawei_{model}.csv"
            prod_bytes = find_definition_file(PROD_DIR, filename).read_bytes()
            assert find_definition_file(REPO_ROOT / "Definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "Definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "generated_definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "tagged_definitions", filename).read_bytes() == prod_bytes
