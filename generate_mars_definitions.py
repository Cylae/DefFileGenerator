#!/usr/bin/env python3
"""Forwarding wrapper for scripts/generate_mars_definitions.py."""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.generate_mars_definitions import (  # noqa: E402
    build_definition_file,
    determine_action,
    generate_all_mars_files,
    main,
    normalize_type,
    normalize_unit,
    parse_mars_workbook,
)

__all__ = [
    "build_definition_file",
    "determine_action",
    "generate_all_mars_files",
    "main",
    "normalize_type",
    "normalize_unit",
    "parse_mars_workbook",
]

if __name__ == "__main__":
    main()

