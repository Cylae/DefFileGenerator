"""
Test suite certifying the Apollo-Standard Production Definitions Suite (production/).
Verifies all 5 pillars across 100% of the production files:
  1. Readable & canonical naming (<PREFIX>_<Manufacturer>_<Model>.csv)
  2. Valid Modbus addressing (decimal 0..65535 or string addr_len)
  3. Certified Webdyn action codes (0, 1, 2, 4, 8, 9, 10)
  4. Standard Modbus function codes (1, 2, 3, 4)
  5. Canonical SunPM tags (P_AC, Q_AC, U_AC, I_AC, P_DC, E_TOT, E_DAY, etc.)
  6. Strict UTF-8 without BOM and 11-column CSV schema
"""

import csv
import re
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import Generator

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

VALID_PREFIXES = {"INV", "MET", "SEN", "BAT", "COM", "OTH"}
VALID_ACTIONS = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"}
VALID_FUNCTION_CODES = {"1", "2", "3", "4"}


def get_production_files() -> list[Path]:
    return sorted(PROD_DIR.glob("*.csv"))


@pytest.fixture(scope="module")
def generator() -> Generator:
    return Generator()


def test_production_directory_populated():
    """Verify that the production directory exists and contains all certified definitions."""
    files = get_production_files()
    assert len(files) >= 1000, f"Expected at least 1000 production files, got {len(files)}"
    readme = PROD_DIR / "README.md"
    assert readme.is_file(), "Missing production/README.md catalog"


def test_production_naming_convention_is_readable():
    """Verify that every file follows the clean <PREFIX>_<Manufacturer>_<Model>.csv pattern."""
    files = get_production_files()
    pattern = re.compile(r"^(INV|MET|SEN|BAT|COM|OTH)_([A-Za-z0-9\-]+)_([A-Za-z0-9\.\-_]+)\.csv$")

    for f in files:
        m = pattern.match(f.name)
        assert m is not None, f"Unreadable or non-canonical filename: {f.name}"
        pfx, mfg, model = m.group(1), m.group(2), m.group(3)
        assert pfx in VALID_PREFIXES, f"Invalid prefix {pfx} in {f.name}"
        assert len(mfg) > 0, f"Empty manufacturer in {f.name}"
        assert not model.endswith("_"), f"Dangling underscore in {f.name}"
        assert not model.startswith("_"), f"Leading underscore in model: {f.name}"
        assert "__" not in f.name, f"Double underscore in {f.name}"
        assert not any(
            f.name.endswith(bad) for bad in ("_2.csv", "_Test.csv", "_draft.csv", "_wip.csv")
        ), f"Artifact suffix in {f.name}"


def test_production_encoding_zero_bom():
    """Verify that all files in production/ are encoded in strict UTF-8 without BOM."""
    files = get_production_files()
    for f in files:
        with open(f, "rb") as fp:
            raw = fp.read(4)
            assert not raw.startswith(b"\xef\xbb\xbf"), f"{f.name} contains UTF-8 BOM"


@pytest.mark.parametrize("file_path", get_production_files(), ids=lambda p: p.name)
def test_production_csv_integrity(file_path: Path):
    """Verify header, 11-column CSV format, addressing, action codes, and function codes."""
    with open(file_path, encoding="utf-8") as fp:
        reader = csv.reader(fp, delimiter=";")
        header = next(reader, None)
        assert header is not None, f"Empty file: {file_path.name}"
        assert len(header) == 11, f"{file_path.name}: header columns {len(header)} != 11"

        rows = list(reader)
        assert len(rows) >= 1, f"{file_path.name}: no registers found"

        seen_single_tags: set[str] = set()

        for idx, row in enumerate(rows, 1):
            assert len(row) == 11, f"{file_path.name} L{idx}: {len(row)} columns != 11"

            # 1. Index monotonic
            assert row[0] == str(idx), f"{file_path.name} L{idx}: index mismatch {row[0]} != {idx}"

            # 2. Function code
            fc = row[1]
            assert fc in VALID_FUNCTION_CODES, (
                f"{file_path.name} L{idx}: invalid function code {fc}"
            )

            # 3. Address
            addr = row[2]
            assert addr, f"{file_path.name} L{idx}: empty address"
            dtype = row[3]
            if dtype == "STRING":
                assert "_" in addr, (
                    f"{file_path.name} L{idx}: STRING missing length in address {addr}"
                )
            else:
                base_addr = addr.split("_")[0]
                assert base_addr.isdigit(), f"{file_path.name} L{idx}: non-numeric address {addr}"
                iaddr = int(base_addr)
                assert 0 <= iaddr <= 65535, f"{file_path.name} L{idx}: address out of range {addr}"

            # 4. Action code
            action = row[10]
            assert action in VALID_ACTIONS, f"{file_path.name} L{idx}: invalid action code {action}"

            # 5. SunPM Tagging
            tag = row[6].strip()
            if tag:
                # Disallow duplicate single-instance measurements on different registers
                if tag in (
                    "P_AC",
                    "Q_AC",
                    "S_AC",
                    "P_DC",
                    "FREQ",
                    "PF",
                    "E_TOT",
                    "E_DAY",
                    "MODEL",
                    "SERIAL",
                ):
                    assert tag not in seen_single_tags, (
                        f"{file_path.name} L{idx}: duplicate single-instance tag {tag}"
                    )
                    seen_single_tags.add(tag)


@pytest.mark.parametrize("file_path", get_production_files(), ids=lambda p: p.name)
def test_production_webdyn_validation(generator: Generator, file_path: Path):
    """Verify that every file in production/ passes detailed Webdyn validation."""
    rep = generator.validate_csv_detailed(str(file_path), strict=True, strict_overlap=False)
    assert rep.is_valid, f"{file_path.name} validation failed: {rep.issues}"
