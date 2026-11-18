#!/usr/bin/env python3
"""Forwarding wrapper for scripts/generate_mars_definitions.py."""
import runpy
from pathlib import Path

if __name__ == "__main__":
    target = Path(__file__).resolve().parent / "scripts" / "generate_mars_definitions.py"
    runpy.run_path(str(target), run_name="__main__")
