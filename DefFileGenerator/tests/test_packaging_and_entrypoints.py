import importlib.metadata
import os
import shutil
import subprocess
import sys
import unittest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


class TestPackagingAndEntrypoints(unittest.TestCase):
    def test_package_metadata(self):
        dist = importlib.metadata.distribution("def-file-generator")
        self.assertEqual(dist.metadata["Version"], "0.2.1")
        self.assertIn("Cylae", dist.metadata["Author"])
        lic = dist.metadata.get("License-Expression") or dist.metadata.get("License", "")
        self.assertEqual(lic, "MIT")

    def test_entrypoint_defined_in_metadata(self):
        entry_points = importlib.metadata.entry_points(group="console_scripts")
        deffilegen_ep = [ep for ep in entry_points if ep.name == "deffilegen"]
        self.assertEqual(len(deffilegen_ep), 1, "Expected 'deffilegen' console script entry point")
        self.assertEqual(deffilegen_ep[0].value, "DefFileGenerator.main:main")

    def test_deffilegen_cli_version_via_module(self):
        res = subprocess.run(
            [sys.executable, "-m", "DefFileGenerator.main", "--version"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("deffilegen 0.2.1", res.stdout)

    def test_deffilegen_cli_help(self):
        res = subprocess.run(
            [sys.executable, "-m", "DefFileGenerator.main", "--help"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("usage: deffilegen", res.stdout)
        self.assertIn("extract", res.stdout)
        self.assertIn("generate", res.stdout)
        self.assertIn("validate", res.stdout)
        self.assertIn("run", res.stdout)

    def test_deffilegen_executable_in_environment(self):
        exe_path = shutil.which("deffilegen")
        if not exe_path:
            # Check virtualenv scripts path directly
            scripts_dir = os.path.dirname(sys.executable)
            possible = os.path.join(scripts_dir, "deffilegen.exe")
            if os.path.exists(possible):
                exe_path = possible
            else:
                possible_unix = os.path.join(scripts_dir, "deffilegen")
                if os.path.exists(possible_unix):
                    exe_path = possible_unix

        self.assertIsNotNone(
            exe_path,
            "deffilegen console script executable should be present in the Python environment",
        )
        if exe_path:
            res = subprocess.run(
                [exe_path, "--version"],
                capture_output=True,
                text=True,
                check=True,
            )
            self.assertIn("deffilegen 0.2.1", res.stdout)


if __name__ == "__main__":
    unittest.main()
