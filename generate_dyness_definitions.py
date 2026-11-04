#!/usr/bin/env python3
"""
CLI entry point for Dyness DH-Series WebdynSunPM definition generator.
Delegates core parsing and generation logic to DefFileGenerator.dyness_generator.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure root directory and .venv site-packages are in sys.path
REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Auto-locate virtual environment site-packages if run with system python in VS Code
for candidate_dir in [REPO_ROOT, REPO_ROOT / "DefFileGenerator"]:
    win_sp = candidate_dir / ".venv" / "Lib" / "site-packages"
    if win_sp.is_dir() and str(win_sp) not in sys.path:
        sys.path.insert(0, str(win_sp))
        break

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
    main()
