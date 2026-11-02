"""Regression: every version surface must agree with pyproject.toml (F-001)."""

import re
import unittest
from pathlib import Path

from fastapi.testclient import TestClient

from DefFileGenerator import __version__
from web.app import app

ROOT = Path(__file__).resolve().parents[2]


class TestVersionConsistency(unittest.TestCase):
    def test_pyproject_matches_package(self):
        text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
        match = re.search(r'^version\s*=\s*"([^"]+)"', text, re.MULTILINE)
        self.assertIsNotNone(match)
        self.assertEqual(match.group(1), __version__)

    def test_web_surfaces_match_package(self):
        self.assertEqual(app.version, __version__)
        res = TestClient(app).get("/api/health")
        self.assertEqual(res.json()["version"], __version__)

    def test_no_stale_hardcoded_version_in_sources(self):
        for rel in ("DefFileGenerator/main.py", "DefFileGenerator/gui.py", "web/app.py"):
            src = (ROOT / rel).read_text(encoding="utf-8")
            self.assertNotRegex(src, r"0\.2\.1", rel)


if __name__ == "__main__":
    unittest.main()
