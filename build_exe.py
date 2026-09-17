#!/usr/bin/env python3
"""
Build orchestration script for WebdynSunPM DefFileGenerator standalone executables.

Compiles the Python package into standalone Windows executables (.exe) using
PyInstaller, performs automated post-build verification and smoke testing,
computes SHA-256 checksums, and optionally packages output into a zip archive.

Usage:
    python build_exe.py [--target {all,gui,cli}] [--clean] [--zip] [--skip-verify]
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess  # nosec: B404
import sys
import zipfile
from pathlib import Path

# Ensure UTF-8 stream output on Windows terminals and CI runners
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Project root directory
REPO_ROOT = Path(__file__).resolve().parent
DIST_DIR = REPO_ROOT / "dist"
BUILD_DIR = REPO_ROOT / "build"
SPEC_FILE = REPO_ROOT / "DefFileGenerator.spec"


def get_project_version() -> str:
    """Extract project version from pyproject.toml without external dependencies."""
    pyproject = REPO_ROOT / "pyproject.toml"
    if pyproject.is_file():
        for line in pyproject.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("version ="):
                return line.split("=")[1].strip().strip('"').strip("'")
    return "0.0.0"


def compute_sha256(filepath: Path) -> str:
    """Calculate the SHA-256 hexadecimal digest of a binary file.

    Args:
        filepath: Absolute or relative path to the file.

    Returns:
        Hexadecimal representation of the SHA-256 hash.
    """
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def clean_artifacts() -> None:
    """Remove previous build and dist directories to ensure a clean build."""
    print("[CLEAN] Removing previous build artifacts...")
    for directory in [BUILD_DIR, DIST_DIR]:
        if directory.is_dir():
            shutil.rmtree(directory, ignore_errors=True)
            print(f"   Removed: {directory}")


def check_pyinstaller() -> None:
    """Ensure PyInstaller is installed and available in the current environment."""
    try:
        import PyInstaller  # noqa: F401 # type: ignore[import-untyped]
    except ImportError:
        print("[ERROR] PyInstaller is not installed in the current Python environment.")
        print('   Please install it via: pip install pyinstaller or pip install -e ".[build]"')
        sys.exit(1)


def run_pyinstaller(target: str) -> None:
    """Execute PyInstaller build using DefFileGenerator.spec.

    Args:
        target: Target executable to compile ('all', 'gui', or 'cli').
    """
    check_pyinstaller()
    print(f"[BUILD] Starting PyInstaller compilation (target: {target})...")

    env = os.environ.copy()
    env["BUILD_TARGET"] = target
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        str(SPEC_FILE),
    ]

    print(f"   Executing: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(REPO_ROOT), env=env)  # nosec: B603
    if result.returncode != 0:
        print(f"[ERROR] PyInstaller build failed with exit code {result.returncode}")
        sys.exit(result.returncode)

    print("[OK] PyInstaller compilation finished successfully.")


def verify_build(target: str) -> bool:
    """Perform post-build verification and smoke testing on generated executables.

    Args:
        target: Target executable that was built ('all', 'gui', or 'cli').

    Returns:
        True if all expected binaries pass verification, False otherwise.
    """
    print("\n[VERIFY] Verifying built executables...")
    all_passed = True

    gui_exe = DIST_DIR / "DefFileGenerator-GUI.exe"
    cli_exe = DIST_DIR / "deffilegen.exe"

    if target in ("all", "gui"):
        if gui_exe.is_file():
            size_mb = gui_exe.stat().st_size / (1024 * 1024)
            print(f"   [OK] [GUI] Found {gui_exe.name} ({size_mb:.2f} MB)")
            if size_mb < 5.0:
                print("   [WARN] GUI executable seems unusually small (< 5 MB).")
        else:
            print(f"   [ERROR] [GUI] Missing expected executable: {gui_exe}")
            all_passed = False

    if target in ("all", "cli"):
        if cli_exe.is_file():
            size_mb = cli_exe.stat().st_size / (1024 * 1024)
            print(f"   [OK] [CLI] Found {cli_exe.name} ({size_mb:.2f} MB)")

            # Smoke test CLI execution on Windows
            if sys.platform == "win32":
                print("   [TEST] Testing CLI execution (deffilegen --version)...")
                try:
                    res = subprocess.run(  # nosec: B603
                        [str(cli_exe), "--version"],
                        capture_output=True,
                        text=True,
                        timeout=15,
                    )
                    if res.returncode == 0 and "deffilegen" in (res.stdout + res.stderr):
                        print(f"      Output: {(res.stdout or res.stderr).strip()}")
                        print("      [OK] CLI smoke test PASSED.")
                    else:
                        print(
                            f"      [ERROR] CLI smoke test failed (code {res.returncode}): {res.stderr}"
                        )
                        all_passed = False
                except Exception as exc:
                    print(f"      [ERROR] CLI smoke test raised exception: {exc}")
                    all_passed = False
        else:
            print(f"   [ERROR] [CLI] Missing expected executable: {cli_exe}")
            all_passed = False

    return all_passed


def generate_checksums() -> Path:
    """Calculate and write SHA-256 checksums for all executables in dist/.

    Returns:
        Path to the generated SHA256SUMS.txt file.
    """
    checksum_file = DIST_DIR / "SHA256SUMS.txt"
    entries: list[str] = []

    print("\n[HASH] Calculating SHA-256 Checksums:")
    for exe in sorted(DIST_DIR.glob("*.exe")):
        sha = compute_sha256(exe)
        line = f"{sha}  {exe.name}"
        entries.append(line)
        print(f"   {exe.name}: {sha}")

    checksum_file.write_text("\n".join(entries) + "\n", encoding="utf-8")
    print(f"   Saved checksums to: {checksum_file}")
    return checksum_file


def create_zip_archive(version: str) -> Path:
    """Bundle executables, documentation, and checksums into a distributable zip file.

    Args:
        version: Project version string.

    Returns:
        Path to the created zip archive.
    """
    archive_name = f"DefFileGenerator-v{version}-windows-x64.zip"
    archive_path = DIST_DIR / archive_name

    print(f"\n[PACK] Packaging release archive: {archive_name}...")
    files_to_pack = list(DIST_DIR.glob("*.exe"))
    checksum_file = DIST_DIR / "SHA256SUMS.txt"
    if checksum_file.is_file():
        files_to_pack.append(checksum_file)

    for doc in ["README.md", "QUICKSTART.md"]:
        doc_path = REPO_ROOT / doc
        if doc_path.is_file():
            files_to_pack.append(doc_path)

    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in files_to_pack:
            zipf.write(file, arcname=file.name)
            print(f"   Added to archive: {file.name}")

    archive_size_mb = archive_path.stat().st_size / (1024 * 1024)
    print(f"[OK] Created archive: {archive_path} ({archive_size_mb:.2f} MB)")
    return archive_path


def main() -> None:
    """CLI entrypoint for build script."""
    parser = argparse.ArgumentParser(
        description="Build standalone Windows executables for DefFileGenerator."
    )
    parser.add_argument(
        "--target",
        choices=["all", "gui", "cli"],
        default="all",
        help="Specify which executable target to build (default: all)",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Do not clean build/ and dist/ directories before building",
    )
    parser.add_argument(
        "--skip-verify",
        action="store_true",
        help="Skip post-build verification and smoke testing",
    )
    parser.add_argument(
        "--zip",
        action="store_true",
        help="Create a distributable zip archive containing the built executables",
    )

    args = parser.parse_args()
    version = get_project_version()
    print(f"=== WebdynSunPM DefFileGenerator v{version} Windows Builder ===")

    if not args.no_clean:
        clean_artifacts()

    run_pyinstaller(args.target)

    if not args.skip_verify:
        passed = verify_build(args.target)
        if not passed:
            print("[ERROR] Build verification failed.")
            sys.exit(1)

    generate_checksums()

    if args.zip:
        create_zip_archive(version)

    print("\n[DONE] Build process completed successfully!")


if __name__ == "__main__":
    main()
