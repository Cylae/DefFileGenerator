"""
Comprehensive validation battery for Mars Renewable Local Controller V7.0
and R&D production definition files (INV_HUAWEI_V4 reference benchmark).
"""

import glob
import os

import pytest

from DefFileGenerator.def_gen import Generator

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BIBLE_DIR = r"G:\My Drive\08092026\DEF_PM_R&D"
HUAWEI_REF = os.path.join(BIBLE_DIR, "INV", "HUAWEI", "INV_HUAWEI_V4.csv")


def test_huawei_v4_reference_conformance():
    """Verify that the benchmark reference file INV_HUAWEI_V4.csv validates flawlessly."""
    if not os.path.exists(HUAWEI_REF):
        pytest.skip(f"Bible reference file not found at {HUAWEI_REF}")

    gen = Generator()
    report = gen.validate_csv_detailed(HUAWEI_REF, strict=True)
    assert report.is_valid, f"Huawei V4 reference failed validation: {report.issues}"
    assert report.register_count == 322
    assert report.stats["errors"] == 0
    assert report.stats["warnings"] == 0


def test_mars_local_controller_slave1_strict_validation():
    """Verify MarsLocalController_V7.0_webdyn.csv (Slave 1) passes strict validation with 0 errors and 0 overlaps."""
    target_csv = os.path.join(WORKSPACE_DIR, "MarsLocalController_V7.0_webdyn.csv")
    assert os.path.exists(target_csv), f"Generated definition missing at {target_csv}"

    gen = Generator()
    report = gen.validate_csv_detailed(target_csv, strict=True, strict_overlap=True)
    assert report.is_valid, (
        f"Mars Local Controller Slave 1 failed validation: {[i.message for i in report.issues]}"
    )
    assert report.register_count >= 10000, f"Expected >10000 registers, got {report.register_count}"
    assert report.stats["errors"] == 0

    # Verify header conformance with Huawei V4 pattern (11 columns)
    with open(target_csv, encoding="utf-8-sig") as f:
        header_line = f.readline().rstrip("\r\n").split(";")
        assert len(header_line) == 11, f"Expected 11 columns in header, got {len(header_line)}"
        assert header_line[0] == "modbusTCP"
        assert header_line[1] == "BESS"
        assert header_line[2] == "Mars Energy"

        # Check sample data rows formatting
        for _ in range(50):
            row = f.readline().rstrip("\r\n").split(";")
            assert len(row) == 11, f"Data row must have 11 columns, got {len(row)}: {row}"
            # Check 6-decimal coefficient formatting
            coef_a = float(row[7])
            coef_b = float(row[8])
            assert f"{coef_a:.6f}" == row[7], f"CoefA not formatted to 6 decimals: {row[7]}"
            assert f"{coef_b:.6f}" == row[8], f"CoefB not formatted to 6 decimals: {row[8]}"


def test_mars_local_controller_ems_slave247():
    """Verify MarsLocalController_V7.0_EMS_Slave247.csv passes strict validation."""
    target_csv = os.path.join(WORKSPACE_DIR, "MarsLocalController_V7.0_EMS_Slave247.csv")
    assert os.path.exists(target_csv), f"File missing: {target_csv}"

    gen = Generator()
    report = gen.validate_csv_detailed(target_csv, strict=True, strict_overlap=True)
    assert report.is_valid, f"Slave 247 failed validation: {report.issues}"
    assert report.register_count == 12

    # Check write action code on EMS commands
    with open(target_csv, encoding="utf-8-sig") as f:
        rows = [line.strip().split(";") for line in f if line.strip()][1:]
        writes = [r for r in rows if r[1] == "3"]
        assert len(writes) == 7
        for w in writes:
            assert w[10] == "1", f"EMS write command must have Action=1, got {w[10]}"


def test_mars_local_controller_rack_slave_n():
    """Verify MarsLocalController_V7.0_RackCell_SlaveN.csv passes strict validation."""
    target_csv = os.path.join(WORKSPACE_DIR, "MarsLocalController_V7.0_RackCell_SlaveN.csv")
    assert os.path.exists(target_csv), f"File missing: {target_csv}"

    gen = Generator()
    report = gen.validate_csv_detailed(target_csv, strict=True, strict_overlap=True)
    assert report.is_valid, f"Rack/Cell Slave N failed: {report.issues}"
    assert report.register_count == 5030


def test_all_134_r_and_d_bible_definitions():
    """Verify 100% of the 134 production definition files in the R&D Bible pass validation."""
    if not os.path.isdir(BIBLE_DIR):
        pytest.skip(f"R&D Bible directory not found at {BIBLE_DIR}")

    csv_files = glob.glob(os.path.join(BIBLE_DIR, "**", "*.csv"), recursive=True)
    assert len(csv_files) >= 130, f"Expected ~134 bible files, found {len(csv_files)}"

    gen = Generator()
    failed = []
    for f in csv_files:
        report = gen.validate_csv_detailed(f, strict=False)
        if not report.is_valid:
            errors = [i.message for i in report.issues if i.severity == "ERROR"]
            failed.append((os.path.basename(f), errors))

    assert len(failed) == 0, f"{len(failed)} files failed in R&D Bible: {failed}"
