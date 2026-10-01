"""
Unit tests for DefFileGenerator.io_utils filesystem safeguards and helper functions.
"""

import os
import tempfile
import unittest

from DefFileGenerator.io_utils import (
    definition_preview,
    paths_refer_to_same_file,
    staged_text_output,
)


class TestIOUtils(unittest.TestCase):
    def test_definition_preview_normal_and_limit(self) -> None:
        """Verifies definition_preview returns correctly mapped rows up to the limit."""
        csv_content = (
            "modbusRTU;Inverter;Vendor;Model;;;;;;;\n"
            "1;1;10001;U16;;Coil Tag 1;coil_tag1;1.0;0.0;;1\n"
            "1;2;10002;U16;;Discrete Tag;disc_tag;1.0;0.0;;2\n"
            "1;3;40001;U16;;Holding Tag;hold_tag;1.0;0.0;W;4\n"
            "1;4;30001;U16;;Input Tag;input_tag;1.0;0.0;V;3\n"
            "1;99;50001;U16;;Custom Tag;custom_tag;1.0;0.0;;4\n"
        )
        preview = definition_preview(csv_content, limit=3)
        self.assertEqual(len(preview), 3)

        self.assertEqual(preview[0]["RegisterType"], "Coils")
        self.assertEqual(preview[0]["Address"], "10001")
        self.assertEqual(preview[0]["Tag"], "coil_tag1")

        self.assertEqual(preview[1]["RegisterType"], "Discrete Inputs")
        self.assertEqual(preview[2]["RegisterType"], "Holding Register")

    def test_definition_preview_truncated_and_empty_rows(self) -> None:
        """Verifies definition_preview handles short/truncated rows and empty lines without error."""
        csv_content = (
            "header\n"
            "1;3\n"  # Truncated row: only 2 fields
            "\n"  # Empty line
            "1;4;30001;U16\n"  # Truncated row: only 4 fields
            "1;1;10001;U16;;Name;Tag;1;0;V;4;extra\n"  # Extra fields
        )
        preview = definition_preview(csv_content, limit=10)
        self.assertEqual(len(preview), 3)

        self.assertEqual(preview[0]["RegisterType"], "Holding Register")
        self.assertEqual(preview[0]["Address"], "")
        self.assertEqual(preview[0]["Action"], "")

        self.assertEqual(preview[1]["RegisterType"], "Input Register")
        self.assertEqual(preview[1]["Address"], "30001")
        self.assertEqual(preview[1]["Name"], "")

        self.assertEqual(preview[2]["RegisterType"], "Coils")
        self.assertEqual(preview[2]["Unit"], "V")
        self.assertEqual(preview[2]["Action"], "4")

    def test_paths_refer_to_same_file(self) -> None:
        """Verifies paths_refer_to_same_file accurately identifies identical files and distinct paths."""
        with tempfile.NamedTemporaryFile(delete=False) as f:
            path1 = f.name

        try:
            path1_relative = os.path.join(os.path.dirname(path1), ".", os.path.basename(path1))
            self.assertTrue(paths_refer_to_same_file(path1, path1_relative))

            path2 = path1 + "_different"
            self.assertFalse(paths_refer_to_same_file(path1, path2))
        finally:
            if os.path.exists(path1):
                os.unlink(path1)

    def test_staged_text_output_atomic_write(self) -> None:
        """Verifies staged_text_output writes atomically and cleans up on error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            target_path = os.path.join(tmpdir, "output.txt")

            with staged_text_output(target_path) as out:
                out.write("Hello Atomic World\n")

            self.assertTrue(os.path.exists(target_path))
            with open(target_path, encoding="utf-8") as f:
                content = f.read()
            self.assertEqual(content, "Hello Atomic World\n")

            # Error handling during staged write
            target_path_err = os.path.join(tmpdir, "failed_output.txt")
            with self.assertRaises(RuntimeError):
                with staged_text_output(target_path_err) as out:
                    out.write("Partial write")
                    raise RuntimeError("Simulated crash during write")

            self.assertFalse(os.path.exists(target_path_err))


if __name__ == "__main__":
    unittest.main()
