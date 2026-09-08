#!/usr/bin/env python3
"""
Single-Step Documentation to WebdynSunPM Converter Interface.

Parses manufacturer register documentation (PDF, Excel, CSV, XML) and generates a
validated WebdynSunPM definition CSV file in a single step.
"""

import argparse
import json
import logging
import os
import re
import json
from DefFileGenerator.main import main as run_cli

def _run_cli(args=None):
    """
    Internal CLI runner that handles argument translation.
    """
    if args is None:
        args = sys.argv[1:]

    commands = ['validate', 'extract', 'generate', 'run']
    has_command = any(arg in commands for arg in args)

    if not has_command and not ('-h' in args or '--help' in args):
        # We need to be careful. If the user passed arguments that look like they
        # belong to 'run' (like --manufacturer), but didn't say 'run', we add 'run'.
        # However, many tests call main() directly with old-style arguments.
        # Let's see what's in args.
        args = ['run'] + args

    # Special case for tests that expect files to be generated in current dir
    # with specific names if -o is not provided.
    run_cli(args)

def main(args=None):
    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit as e:
        sys.exit(e.code)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        import traceback
        logging.debug(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()