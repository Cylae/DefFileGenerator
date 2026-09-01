"""Guards reproducible output across interpreter hash-seed randomisation."""

import hashlib
import os
import subprocess
import sys
import tempfile
import unittest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
CLI = os.path.join(REPO_ROOT, 'DefFileGenerator', 'main.py')


class TestDeterministicOutput(unittest.TestCase):
    """Column detection once iterated a set, so binding varied per process."""

    FIXTURES = ('template.csv', 'sample_inverter_registers.xlsx', 'type_synonyms.csv')

    def _digest(self, fixture, seed):
        fd, target = tempfile.mkstemp(suffix='.csv')
        os.close(fd)
        os.unlink(target)
        env = dict(os.environ, PYTHONPATH=REPO_ROOT, PYTHONHASHSEED=str(seed))
        try:
            subprocess.run(
                [sys.executable, CLI, 'run', fixture, '--manufacturer', 'A',
                 '--model', 'B', '--address-offset', '250', '-o', target],
                capture_output=True, env=env, cwd=REPO_ROOT,
            )
            if not os.path.exists(target):
                return None
            with open(target, 'rb') as f:
                return hashlib.md5(f.read()).hexdigest()
        finally:
            if os.path.exists(target):
                os.unlink(target)

    def test_output_is_stable_across_hash_seeds(self):
        for fixture in self.FIXTURES:
            if not os.path.exists(os.path.join(REPO_ROOT, fixture)):
                continue
            with self.subTest(fixture=fixture):
                digests = {self._digest(fixture, seed) for seed in range(4)}
                self.assertEqual(
                    len(digests), 1,
                    f"{fixture} produced {len(digests)} different outputs "
                    f"across PYTHONHASHSEED values",
                )

    def test_compound_string_address_survives_offset(self):
        """The STR<n> row was dropped on seeds where Address bound to the wrong column."""
        fd, target = tempfile.mkstemp(suffix='.csv')
        os.close(fd)
        os.unlink(target)
        env = dict(os.environ, PYTHONPATH=REPO_ROOT, PYTHONHASHSEED='1')
        try:
            subprocess.run(
                [sys.executable, CLI, 'run', 'template.csv', '--manufacturer', 'A',
                 '--model', 'B', '--address-offset', '250', '-o', target],
                capture_output=True, env=env, cwd=REPO_ROOT,
            )
            with open(target, encoding='utf-8-sig') as f:
                content = f.read()
            self.assertIn('30251', content)
            self.assertIn('30280_20', content)
        finally:
            if os.path.exists(target):
                os.unlink(target)


if __name__ == '__main__':
    unittest.main()
