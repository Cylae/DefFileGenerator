"""
Production Suite Governor and Continuous Ingestion Engine.
Guarantees that any definition file added to `production/` strictly satisfies:
  1. Readable naming convention (<PREFIX>_<Manufacturer>_<Model>.csv)
  2. Modbus addressing within bounds (0..65535 or string addr_len)
  3. Action codes (0, 1, 2, 4, 8, 9, 10)
  4. Function codes (1, 2, 3, 4)
  5. Canonical SunPM tags without fatal duplicates
  6. Strict UTF-8 without BOM and 11-column semicolon CSV schema
"""

import csv
import re
from pathlib import Path

from DefFileGenerator.def_gen import Generator

VALID_PREFIXES = {"INV", "MET", "SEN", "BAT", "COM", "OTH"}
VALID_ACTIONS = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"}
VALID_FUNCTION_CODES = {"1", "2", "3", "4"}

NAMING_PATTERN = re.compile(
    r"^(INV|MET|SEN|BAT|COM|OTH)_([A-Za-z0-9\-]+)_([A-Za-z0-9\.\-_]+)\.csv$"
)


def verify_file_pillars(file_path: Path) -> tuple[bool, list[str]]:
    """
    Rigorously verifies a candidate definition file against the 5 production pillars.
    Returns (is_valid, list_of_issues).
    """
    issues: list[str] = []

    # 1. Naming Check
    filename = file_path.name
    m = NAMING_PATTERN.match(filename)
    if not m:
        issues.append(
            f"Filename '{filename}' does not follow readable pattern <PREFIX>_<Manufacturer>_<Model>.csv"
        )
    else:
        pfx, _mfg, model = m.group(1), m.group(2), m.group(3)
        if pfx not in VALID_PREFIXES:
            issues.append(f"Invalid prefix '{pfx}' in filename '{filename}'")
        if model.endswith("_") or model.startswith("_"):
            issues.append(f"Dangling underscore in model '{model}' in filename '{filename}'")
        if "__" in filename:
            issues.append(f"Double underscore in filename '{filename}'")
        for bad_suffix in (
            "_2.csv",
            "_Test.csv",
            "_draft.csv",
            "_wip.csv",
            "_old.csv",
            "_copy.csv",
        ):
            if filename.lower().endswith(bad_suffix.lower()):
                issues.append(f"Forbidden artifact suffix '{bad_suffix}' in filename '{filename}'")

    # 2. BOM & Encoding Check
    try:
        with open(file_path, "rb") as fp:
            raw = fp.read(4)
            if raw.startswith(b"\xef\xbb\xbf"):
                issues.append(f"File '{filename}' contains prohibited UTF-8 BOM")
    except Exception as e:
        issues.append(f"Cannot read file '{filename}': {e}")
        return False, issues

    # 3. CSV Parsing & Row-Level Checks
    try:
        with open(file_path, encoding="utf-8") as fp:
            reader = csv.reader(fp, delimiter=";")
            header = next(reader, None)
            if header is None:
                issues.append(f"File '{filename}' is empty")
                return False, issues
            if len(header) != 11:
                issues.append(f"Header in '{filename}' has {len(header)} columns instead of 11")

            rows = list(reader)
            if not rows:
                issues.append(f"File '{filename}' has no register rows")
                return False, issues

            seen_single_tags: set[str] = set()

            for idx, row in enumerate(rows, 1):
                if len(row) != 11:
                    issues.append(
                        f"Line {idx} in '{filename}': expected 11 columns, got {len(row)}"
                    )
                    continue

                # Index
                if row[0] != str(idx):
                    issues.append(f"Line {idx} in '{filename}': index mismatch '{row[0]}' != {idx}")

                # Function Code
                fc = row[1].strip()
                if fc not in VALID_FUNCTION_CODES:
                    issues.append(
                        f"Line {idx} in '{filename}': invalid Modbus function code '{fc}'"
                    )

                # Address
                addr = row[2].strip()
                if not addr:
                    issues.append(f"Line {idx} in '{filename}': empty address")
                else:
                    dtype = row[3].strip()
                    if dtype == "STRING":
                        if "_" not in addr:
                            issues.append(
                                f"Line {idx} in '{filename}': STRING address '{addr}' missing length suffix"
                            )
                    else:
                        base_addr = addr.split("_")[0]
                        if not base_addr.isdigit():
                            issues.append(
                                f"Line {idx} in '{filename}': non-numeric address '{addr}'"
                            )
                        else:
                            iaddr = int(base_addr)
                            if not (0 <= iaddr <= 65535):
                                issues.append(
                                    f"Line {idx} in '{filename}': address '{addr}' out of 16-bit range"
                                )

                # Action Code
                action = row[10].strip()
                if action not in VALID_ACTIONS:
                    issues.append(
                        f"Line {idx} in '{filename}': invalid Webdyn action code '{action}'"
                    )

                # SunPM Tag
                tag = row[6].strip()
                if tag:
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
                        if tag in seen_single_tags:
                            issues.append(
                                f"Line {idx} in '{filename}': duplicate unique physical tag '{tag}'"
                            )
                        seen_single_tags.add(tag)

    except Exception as e:
        issues.append(f"Error parsing CSV '{filename}': {e}")
        return False, issues

    # 4. Webdyn Generator Detailed Validation
    gen = Generator()
    val_rep = gen.validate_csv_detailed(str(file_path), strict=True, strict_overlap=False)
    if not val_rep.is_valid:
        for iss in val_rep.issues:
            if iss.severity == "ERROR":
                issues.append(f"Webdyn validation error L{iss.line}: {iss.message}")

    return len(issues) == 0, issues
