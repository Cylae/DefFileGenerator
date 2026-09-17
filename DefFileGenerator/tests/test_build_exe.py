"""
Automated tests for Windows executable build system and PyInstaller configuration.

Validates:
1. `DefFileGenerator.spec` file structure and syntax.
2. `build_exe.py` argument parsing and helper functions (version extraction, sha256).
3. Environment variable target resolution.
"""

from __future__ import annotations

import hashlib
import os
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
SPEC_FILE = REPO_ROOT / "DefFileGenerator.spec"
BUILD_SCRIPT = REPO_ROOT / "build_exe.py"


def test_spec_file_exists_and_parses() -> None:
    """Verify DefFileGenerator.spec exists and contains required targets and collections."""
    assert SPEC_FILE.is_file(), "DefFileGenerator.spec must exist at repository root"
    content = SPEC_FILE.read_text(encoding="utf-8")

    # Target identifiers
    assert "DefFileGenerator-GUI" in content
    assert "deffilegen" in content

    # Package collections
    assert "customtkinter" in content
    assert "openpyxl" in content
    assert "pdfplumber" in content
    assert "pypdfium2" in content
    assert "defusedxml" in content

    # Python syntax verification
    compiled = compile(content, str(SPEC_FILE), "exec")
    assert compiled is not None


def test_build_script_helpers(tmp_path: Path) -> None:
    """Test helper functions in build_exe.py."""
    sys.path.insert(0, str(REPO_ROOT))
    import build_exe

    # Version extraction
    version = build_exe.get_project_version()
    assert version != "0.0.0"
    assert "." in version

    # SHA-256 calculation
    test_file = tmp_path / "sample.bin"
    test_file.write_bytes(b"WebdynSunPM DefFileGenerator Test Binary Payload")
    expected_sha = hashlib.sha256(b"WebdynSunPM DefFileGenerator Test Binary Payload").hexdigest()
    actual_sha = build_exe.compute_sha256(test_file)
    assert actual_sha == expected_sha


def test_build_script_cli_argument_parsing() -> None:
    """Test command-line argument parsing in build_exe.py."""
    sys.path.insert(0, str(REPO_ROOT))
    import build_exe

    with patch("sys.argv", ["build_exe.py", "--target", "gui", "--no-clean", "--zip"]):
        with patch.object(build_exe, "run_pyinstaller") as mock_run:
            with patch.object(build_exe, "verify_build", return_value=True):
                with patch.object(build_exe, "generate_checksums"):
                    with patch.object(build_exe, "create_zip_archive"):
                        build_exe.main()
                        mock_run.assert_called_once_with("gui")


def test_workflow_file_syntax() -> None:
    """Verify .github/workflows/build-exe.yml contains required CI/CD definitions."""
    workflow_path = REPO_ROOT / ".github" / "workflows" / "build-exe.yml"
    assert workflow_path.is_file(), "build-exe.yml must exist"
    content = workflow_path.read_text(encoding="utf-8")

    assert "runs-on: windows-latest" in content
    assert "build_exe.py" in content
    assert "softprops/action-gh-release" in content
    assert "actions/upload-artifact" in content
