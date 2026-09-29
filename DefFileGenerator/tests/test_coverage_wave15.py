import logging
import os
import unittest
from unittest.mock import patch

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor


class TestCoverageWave15(unittest.TestCase):
    def test_apply_address_offset_edge_cases(self):
        """Test Generator.apply_address_offset with invalid/unusual inputs."""
        # Non-numeric string address
        res = Generator.apply_address_offset("invalid_addr", 10)
        self.assertEqual(res, "invalid_addr")

        # Negative offset resulting in negative address warning
        with self.assertLogs(level=logging.WARNING) as cm:
            res_neg = Generator.apply_address_offset("5", -10)
            self.assertEqual(res_neg, "-5")
            self.assertTrue(any("negative address" in log.lower() for log in cm.output))

        # Hex address with underscore compound address
        res_hex_compound = Generator.apply_address_offset("0x10_0x20", 5)
        self.assertEqual(res_hex_compound, "21_32")

    def test_calculate_coefficients_extreme_floats(self):
        """Test Generator._calculate_coefficients with extreme float/non-finite values."""
        # NaN / Infinity inputs with scale_factor
        a, b = Generator._calculate_coefficients("NaN", "0", "1")
        self.assertEqual((a, b), ("10.000000", "0.000000"))

        a_inf, b_inf = Generator._calculate_coefficients("inf", "0", "1")
        self.assertEqual((a_inf, b_inf), ("10.000000", "0.000000"))

    def test_extractor_pdf_page_range_parsing(self):
        """Test Extractor handling of PDF page range arguments."""
        extractor = Extractor()
        # Non-existent file with page range string
        gen = list(extractor.extract_from_pdf("non_existent_file.pdf", pages="5-3, 1, 8-10"))
        self.assertEqual(gen, [])


class TestGuiWave15(unittest.TestCase):
    def test_gui_helpers_without_display(self):
        """Test GUI helper functions and queue logging."""
        import queue

        from DefFileGenerator.gui import QueueLogHandler, main

        q: queue.Queue[tuple[str, str]] = queue.Queue()
        handler = QueueLogHandler(q)
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="test message",
            args=(),
            exc_info=None,
        )
        handler.emit(record)
        self.assertFalse(q.empty())
        level, msg = q.get()
        self.assertEqual(level, "INFO")
        self.assertIn("test message", msg)

        # CLI help and version flags for main()
        with patch.object(os, "_exit"):
            with patch("sys.argv", ["deffilegen-gui", "--version"]):
                main()
            with patch("sys.argv", ["deffilegen-gui", "--help"]):
                main()


if __name__ == "__main__":
    unittest.main()
