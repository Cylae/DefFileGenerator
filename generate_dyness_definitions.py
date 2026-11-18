#!/usr/bin/env python3
"""Forwarding wrapper for scripts/generate_dyness_definitions.py."""

import runpy
import sys
from pathlib import Path

# Ensure root directory is in sys.path
REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from DefFileGenerator.dyness_generator import (  # noqa: E402
    DynessDefinitionGenerator,
    DynessProtocolParser,
    clean_text,
    generate_dyness_definitions,
    main,
    parse_coeff_and_unit,
)

__all__ = [
    "DynessDefinitionGenerator",
    "DynessProtocolParser",
    "clean_text",
    "generate_dyness_definitions",
    "main",
    "parse_coeff_and_unit",
]

if __name__ == "__main__":
    target = REPO_ROOT / "scripts" / "generate_dyness_definitions.py"
    runpy.run_path(str(target), run_name="__main__")
