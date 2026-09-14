#!/usr/bin/env python3
"""
Equipementiers Real-World Battery.

Discovers and tests real-world manufacturer documentation files (XLSX, CSV, XML, PDF)
from the 'G:\\My Drive\\08092026\\Equipementiers' directory, verifying extractor robustness,
field normalization, definition generation, and CSV validation across diverse industry formats.
"""

import logging
import os
import sys
import tempfile
import time
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator

EQUIPEMENTIERS_DIR = r"G:\My Drive\08092026\Equipementiers"


def run_battery(max_files_per_type: int = 15) -> bool:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logger = logging.getLogger("EquipementiersBattery")

    if not os.path.exists(EQUIPEMENTIERS_DIR):
        logger.warning(f"Equipementiers directory not found at: {EQUIPEMENTIERS_DIR}. Skipping.")
        return True

    logger.info(f"Starting Equipementiers Real-World Battery on: {EQUIPEMENTIERS_DIR}")

    extractor = Extractor()
    generator = Generator()
    temp_dir = tempfile.TemporaryDirectory()

    stats = {
        "xlsx": {"tested": 0, "extracted": 0, "registers": 0, "valid_defs": 0, "errors": 0},
        "csv": {"tested": 0, "extracted": 0, "registers": 0, "valid_defs": 0, "errors": 0},
        "xml": {"tested": 0, "extracted": 0, "registers": 0, "valid_defs": 0, "errors": 0},
    }

    start_time = time.time()

    # Discover candidate files
    candidates: dict[str, list[str]] = {"xlsx": [], "csv": [], "xml": []}
    for root, _, files in os.walk(EQUIPEMENTIERS_DIR):
        for f in files:
            ext = os.path.splitext(f)[1].lower().lstrip(".")
            if ext in candidates and len(candidates[ext]) < max_files_per_type * 3:
                full_path = os.path.join(root, f)
                # Skip massive files (> 20MB) to keep battery responsive
                try:
                    if os.path.getsize(full_path) < 20 * 1024 * 1024:
                        candidates[ext].append(full_path)
                except OSError:
                    continue

    for ftype, file_list in candidates.items():
        logger.info(
            f"Discovered {len(file_list)} candidate .{ftype} files. Testing top {min(len(file_list), max_files_per_type)}..."
        )
        for filepath in file_list[:max_files_per_type]:
            stats[ftype]["tested"] += 1
            rel_name = os.path.relpath(filepath, EQUIPEMENTIERS_DIR)
            try:
                if ftype == "xlsx":
                    raw = extractor.extract_from_excel(filepath)
                elif ftype == "csv":
                    raw = extractor.extract_from_csv(filepath)
                elif ftype == "xml":
                    raw = extractor.extract_from_xml(filepath)
                else:
                    continue

                has_data, raw_iter = peek_generator(raw)
                if not has_data:
                    continue

                stats[ftype]["extracted"] += 1
                mapped = list(extractor.map_and_clean(raw_iter))
                reg_count = len(mapped)
                stats[ftype]["registers"] += reg_count

                if reg_count > 0:
                    # Generate Webdyn definition file
                    mfg = Path(filepath).parts[-2] if len(Path(filepath).parts) > 1 else "Mfg"
                    out_path = os.path.join(
                        temp_dir.name, f"def_{stats[ftype]['tested']}_{ftype}.csv"
                    )
                    cfg = GeneratorConfig(
                        input_file=filepath,
                        output=out_path,
                        manufacturer=mfg,
                        model="Model",
                        protocol="modbusRTU",
                        category="Inverter",
                    )
                    run_generator(cfg, input_data=mapped)

                    if os.path.exists(out_path):
                        if generator.validate_csv(out_path, strict=False):
                            stats[ftype]["valid_defs"] += 1
                        os.remove(out_path)

            except Exception as e:
                stats[ftype]["errors"] += 1
                logger.error(f"Error testing {rel_name}: {e}")

    temp_dir.cleanup()
    elapsed = time.time() - start_time

    logger.info("=" * 60)
    logger.info("EQUIPEMENTIERS REAL-WORLD BATTERY SUMMARY")
    logger.info(f"Elapsed Time: {elapsed:.2f}s")
    for ftype, data in stats.items():
        logger.info(
            f"  .{ftype.upper()}: Tested {data['tested']}, Extracted Tables {data['extracted']}, "
            f"Registers {data['registers']}, Valid Definitions {data['valid_defs']}, Errors {data['errors']}"
        )
    logger.info("=" * 60)

    total_errors = sum(s["errors"] for s in stats.values())
    return total_errors == 0


if __name__ == "__main__":
    success = run_battery()
    sys.exit(0 if success else 1)
