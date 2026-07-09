#!/usr/bin/env python3
import sys
import os

# Ensure the package is discoverable
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from DefFileGenerator.main import main

def _run_cli():
    # This is needed by some tests that patch doc_to_webdyn._run_cli
    from DefFileGenerator.main import _run_cli as main_run_cli
    main_run_cli()

if __name__ == "__main__":
    main()
