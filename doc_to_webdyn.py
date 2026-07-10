#!/usr/bin/env python3
import sys
import os
import logging
import argparse

# Ensure the package is discoverable
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '.')))

from DefFileGenerator.main import main as run_cli

def _run_cli(args_list):
    """Internal helper for testing and delegation."""
    run_cli(args_list)

def main(args=None):
    """
    Main entry point for doc_to_webdyn.
    Delegates to the unified CLI in DefFileGenerator.main.
    """
    if args is None:
        args = sys.argv[1:]

    # If no command is provided, default to 'run' for backward compatibility
    # but only if input_file is provided and doesn't match any sub-command name.
    if args and args[0] not in ['extract', 'generate', 'run', 'validate', '-h', '--help']:
        args = ['run'] + args

    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
