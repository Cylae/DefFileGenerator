#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.main import main as cli_main

def _run_cli():
    # This is a helper for unit tests that expect this attribute
    args = sys.argv[1:]
    if not args or args[0] not in ['extract', 'generate', 'run', 'validate']:
        args = ['run'] + args

    # Handle the default output filename logic that tests might expect
    # if -o is not provided and it's a run/generate command
    if ('run' in args or 'generate' in args) and '-o' not in args and '--output' not in args:
        mfg = "Manufacturer"
        model = "Model"
        for i, arg in enumerate(args):
            if arg == '--manufacturer' and i+1 < len(args):
                mfg = args[i+1]
            elif arg == '--model' and i+1 < len(args):
                model = args[i+1]

        output_file = f"{re.sub(r'[^a-zA-Z0-9]', '_', mfg).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', model).lower()}_definition.csv"
        args.extend(['-o', output_file])

    cli_main(args)

def main(args=None):
    try:
        if args is not None:
            # If args are passed programmatically, we need to mock sys.argv for _run_cli
            import unittest.mock
            with unittest.mock.patch.object(sys, 'argv', ['doc_to_webdyn.py'] + args):
                _run_cli()
        else:
            _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
