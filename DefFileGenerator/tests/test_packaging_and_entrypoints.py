import importlib.metadata
import os
import shutil
import site
import subprocess
import sys
import unittest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def _locate_script(name: str) -> str | None:
    """Find a console script by name across all relevant script directories."""
    # 1. Try PATH first (works in venvs and properly installed envs).
    found = shutil.which(name)
    if found:
        return found

    candidates = []

    # 2. System Python scripts dir (same level as sys.executable).
    candidates.append(os.path.dirname(sys.executable))

    # 3. User-base scripts dir (pip install --user on Windows puts scripts here).
    try:
        user_base = site.getuserbase()
        # Windows: <userbase>\PythonXY\Scripts
        ver = f"Python{sys.version_info.major}{sys.version_info.minor}"
        candidates.append(os.path.join(user_base, ver, "Scripts"))
        # Unix: <userbase>/bin
        candidates.append(os.path.join(user_base, "bin"))
    except AttributeError:
        pass

    for scripts_dir in candidates:
        for candidate in (
            os.path.join(scripts_dir, f"{name}.exe"),
            os.path.join(scripts_dir, name),
        ):
            if os.path.exists(candidate):
                return candidate

    return None


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
        exe_path = _locate_script("deffilegen")

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
