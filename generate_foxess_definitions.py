#!/usr/bin/env python3
"""Forwarding wrapper for scripts/generate_foxess_definitions.py."""

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.generate_foxess_definitions import (  # noqa: E402
    build_all_in_one_rows,
    build_battery_rows,
    build_inverter_rows,
    build_meter_rows,
    format_row,
    generate_other_fox_sources,
    main,
    write_definition_file,
)

__all__ = [
    "build_all_in_one_rows",
    "build_battery_rows",
    "build_inverter_rows",
    "build_meter_rows",
    "format_row",
    "generate_other_fox_sources",
    "main",
    "write_definition_file",
]

if __name__ == "__main__":
    main()
