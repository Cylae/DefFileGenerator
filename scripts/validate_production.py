#!/usr/bin/env python3
"""
CLI Tool: Validate Definition Files against Production 5-Pillar Standards.

Usage:
    python scripts/validate_production.py [file_or_directory]
    Default: validates the entire `production/` directory.
"""

import sys
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.production_governor import verify_file_pillars  # noqa: E402
from DefFileGenerator.webdynsun_ini import verify_webdynsun_ini  # noqa: E402


def main():
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else WORKSPACE_DIR / "production"

    if not target.exists():
        print(f"ERROR: Target '{target}' does not exist.")
        sys.exit(1)

    files_to_check: list[Path] = []
    if target.is_file():
        files_to_check.append(target)
    else:
        files_to_check = sorted(
            p for p in target.rglob("*.csv") if "sources" not in p.parts
        )

    print("=" * 80)
    print(f"APOLLO PRODUCTION 5-PILLAR VERIFIER: {len(files_to_check)} file(s)")
    print("=" * 80)

    passed_count = 0
    failed_files: list[tuple[Path, list[str]]] = []

    if not files_to_check:
        print(f"ERROR: No definition CSV files found under '{target}'.")
        sys.exit(1)

    for idx, f in enumerate(files_to_check, 1):
        is_valid, issues = verify_file_pillars(f)
        if target.is_dir():
            ini_file = f.with_suffix(".ini")
            if not ini_file.is_file():
                is_valid = False
                issues.append(f"Missing companion WebdynSun .ini file: {ini_file.name}")
            else:
                ini_ok, ini_issues = verify_webdynsun_ini(ini_file)
                if not ini_ok:
                    is_valid = False
                    issues.extend(ini_issues)
            sources_dir = f.parent / "sources"
            if not sources_dir.is_dir():
                is_valid = False
                issues.append(f"Missing range sources/ directory: {sources_dir}")
        if is_valid:
            passed_count += 1
            print(f"[{idx:03d}/{len(files_to_check)}] PASS: {f.name}")
        else:
            failed_files.append((f, issues))
            print(f"[{idx:03d}/{len(files_to_check)}] FAIL: {f.name}")
            for iss in issues:
                print(f"    -> {iss}")

    print("\n" + "=" * 80)
    print(f"Passed : {passed_count} / {len(files_to_check)}")
    print(f"Failed : {len(failed_files)} / {len(files_to_check)}")
    print("=" * 80)

    if failed_files:
        sys.exit(1)
    else:
        print("ALL PRODUCTION FILES COMPLY WITH THE 5 PILLARS (PASS 100%)")


if __name__ == "__main__":
    main()
