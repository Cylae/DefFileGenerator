"""Regression tests for command-line ergonomics and safety behaviours."""

import os
import subprocess
import sys
import tempfile
import unittest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
CLI = os.path.join(REPO_ROOT, 'DefFileGenerator', 'main.py')
FIXTURE = os.path.join(REPO_ROOT, 'sample_inverter_registers.xlsx')


def run_cli(*args):
    env = dict(os.environ, PYTHONPATH=REPO_ROOT)
    return subprocess.run(
        [sys.executable, CLI, *args],
        capture_output=True, text=True, env=env, cwd=REPO_ROOT,
    )


class TestCliErgonomics(unittest.TestCase):
    def setUp(self):
        fd, self.target = tempfile.mkstemp(suffix='.csv')
        os.close(fd)
        os.unlink(self.target)

    def tearDown(self):
        if os.path.exists(self.target):
            os.unlink(self.target)

    def test_version_flag(self):
        result = run_cli('--version')
        self.assertEqual(result.returncode, 0)
        self.assertIn('deffilegen', result.stdout)

    def test_no_command_is_usage_error(self):
        """Invoking with no sub-command must signal misuse, not success."""
        self.assertEqual(run_cli().returncode, 2)

    def test_unsupported_extension_lists_supported_formats(self):
        fd, bogus = tempfile.mkstemp(suffix='.docx')
        os.close(fd)
        try:
            result = run_cli('run', bogus, '--manufacturer', 'A', '--model', 'B',
                             '-o', self.target)
            self.assertEqual(result.returncode, 1)
            self.assertIn('Unsupported file type', result.stderr)
            self.assertIn('.pdf', result.stderr)
        finally:
            os.unlink(bogus)

    def test_missing_required_flags_names_them(self):
        result = run_cli('run', FIXTURE)
        self.assertEqual(result.returncode, 1)
        self.assertIn('--manufacturer', result.stderr)
        self.assertIn('--model', result.stderr)

    def test_missing_input_file_is_reported(self):
        result = run_cli('run', '/nonexistent/path.csv', '--manufacturer', 'A',
                         '--model', 'B', '-o', self.target)
        self.assertEqual(result.returncode, 1)
        self.assertIn('not found', result.stderr)

    def test_existing_output_is_not_silently_overwritten(self):
        with open(self.target, 'w', encoding='utf-8') as f:
            f.write('SENTINEL\n')
        result = run_cli('run', FIXTURE, '--manufacturer', 'A', '--model', 'B',
                         '-o', self.target)
        self.assertEqual(result.returncode, 1)
        self.assertIn('--force', result.stderr)
        with open(self.target, encoding='utf-8') as f:
            self.assertEqual(f.read().strip(), 'SENTINEL')

    def test_force_allows_overwrite(self):
        with open(self.target, 'w', encoding='utf-8') as f:
            f.write('SENTINEL\n')
        result = run_cli('run', FIXTURE, '--manufacturer', 'A', '--model', 'B',
                         '-o', self.target, '--force')
        self.assertEqual(result.returncode, 0)
        with open(self.target, encoding='utf-8-sig') as f:
            self.assertIn('modbusRTU', f.read())

    def test_verbosity_flags_accepted_after_subcommand(self):
        """Operators habitually append -v/-q after the sub-command."""
        for flag in ('-v', '-q', '--quiet', '--verbose'):
            with self.subTest(flag=flag):
                if os.path.exists(self.target):
                    os.unlink(self.target)
                result = run_cli('run', FIXTURE, '--manufacturer', 'A',
                                 '--model', 'B', '-o', self.target, flag)
                self.assertEqual(result.returncode, 0, result.stderr)

    def test_quiet_suppresses_informational_output(self):
        result = run_cli('run', FIXTURE, '--manufacturer', 'A', '--model', 'B',
                         '-o', self.target, '--quiet')
        self.assertEqual(result.returncode, 0)
        self.assertNotIn('INFO', result.stderr)

    def test_run_validates_generated_file(self):
        result = run_cli('run', FIXTURE, '--manufacturer', 'A', '--model', 'B',
                         '-o', self.target)
        self.assertEqual(result.returncode, 0)
        self.assertIn('Post-generation validation passed', result.stderr)

    def test_no_validate_skips_the_self_check(self):
        result = run_cli('run', FIXTURE, '--manufacturer', 'A', '--model', 'B',
                         '-o', self.target, '--no-validate')
        self.assertEqual(result.returncode, 0)
        self.assertNotIn('Post-generation validation', result.stderr)

    def test_validate_lenient_downgrades_overlaps(self):
        overlapping = os.path.join(REPO_ROOT, 'bad_def.csv')
        if not os.path.exists(overlapping):
            self.skipTest('bad_def.csv fixture unavailable')
        self.assertEqual(run_cli('validate', overlapping, '--lenient').returncode, 0)

    def test_validate_missing_file_is_reported(self):
        result = run_cli('validate', '/nonexistent/def.csv')
        self.assertEqual(result.returncode, 1)
        self.assertIn('not found', result.stderr)

    def test_help_lists_examples(self):
        result = run_cli('--help')
        self.assertEqual(result.returncode, 0)
        self.assertIn('Examples:', result.stdout)
        self.assertIn('Exit codes:', result.stdout)


if __name__ == '__main__':
    unittest.main()
