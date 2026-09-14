"""
Unit and integration tests for Windows 11 Desktop Application (DefFileGenerator.gui).
"""

import logging
import os
import queue
import tempfile
import unittest
from unittest.mock import patch

import DefFileGenerator.gui as gui_module
from DefFileGenerator.def_gen import ValidationIssue, ValidationReport
from DefFileGenerator.gui import (
    EQUIPMENT_TEMPLATES,
    DefFileGenApp,
    QueueLogHandler,
)


class TestGUIComponents(unittest.TestCase):
    def test_equipment_templates_structure(self):
        """Verifies all equipment templates have valid registers and descriptions."""
        self.assertIn("Inverter", EQUIPMENT_TEMPLATES)
        self.assertIn("Meter", EQUIPMENT_TEMPLATES)
        self.assertIn("Sensor", EQUIPMENT_TEMPLATES)
        self.assertIn("Battery", EQUIPMENT_TEMPLATES)
        self.assertIn("WeatherStation", EQUIPMENT_TEMPLATES)
        self.assertIn("Generic", EQUIPMENT_TEMPLATES)

        for name, data in EQUIPMENT_TEMPLATES.items():
            self.assertIn("description", data)
            self.assertIn("registers", data)
            self.assertGreater(len(data["registers"]), 0)
            for reg in data["registers"]:
                # Each register tuple must have 11 fields
                self.assertEqual(
                    len(reg), 11, f"Template {name} register {reg} has invalid column count"
                )

    def test_queue_log_handler(self):
        """Verifies thread-safe log handler pushes records into the queue."""
        log_q: queue.Queue[tuple[str, str]] = queue.Queue()
        handler = QueueLogHandler(log_q)
        record = logging.LogRecord(
            name="test_logger",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="Hello Windows 11",
            args=(),
            exc_info=None,
        )
        handler.emit(record)
        level, msg = log_q.get_nowait()
        self.assertEqual(level, "INFO")
        self.assertIn("Hello Windows 11", msg)


class TestDefFileGenApp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Create a hidden Tk instance for testing UI controllers without popping windows
        try:
            if gui_module.HAS_CUSTOMTKINTER:
                cls.root = gui_module.ctk.CTk()
            else:
                cls.root = gui_module.tk.Tk()
            cls.root.withdraw()
            cls.app = DefFileGenApp(cls.root)
        except Exception as e:
            raise unittest.SkipTest(f"GUI environment not available: {e}") from e

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def test_app_initialization(self):
        """Checks app state and UI controls are properly instantiated."""
        self.assertIsNotNone(self.app.entry_input_file)
        self.assertIsNotNone(self.app.entry_mfg)
        self.assertIsNotNone(self.app.entry_model)
        self.assertIsNotNone(self.app.tree_preview)
        self.assertIsNotNone(self.app.tree_issues)
        self.assertFalse(self.app.is_processing)

    def test_on_conversion_success_callback(self):
        """Tests UI update on conversion completion."""
        mock_rows = [
            {
                "RegisterType": "Holding",
                "Address": "40001",
                "Type": "U16",
                "Name": "Active Power",
                "Tag": "active_power",
                "Factor": "1",
                "Offset": "0",
                "Unit": "W",
                "Action": "4",
            }
        ]
        mock_report = ValidationReport(
            is_valid=True,
            register_count=1,
            issues=[],
            stats={"errors": 0, "warnings": 0, "types": {"3": 1}, "registers": 1},
        )
        with patch.object(gui_module.messagebox, "showinfo") as mock_box:
            self.app._on_conversion_success("test_out.csv", mock_rows, mock_report)
            mock_box.assert_called_once()
            self.assertEqual(self.app.last_generated_file, "test_out.csv")
            self.assertEqual(len(self.app.tree_preview.get_children()), 1)

    def test_on_conversion_error_callback(self):
        """Tests UI update on conversion error."""
        with patch.object(gui_module.messagebox, "showerror") as mock_box:
            self.app._on_conversion_error("Extraction failed badly")
            mock_box.assert_called_once()
            self.assertFalse(self.app.is_processing)

    def test_run_validation_valid_file(self):
        """Tests validator tab execution on a valid file."""
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;U16;;Active Power;active_power;1.0;0.0;W;4\n")
            filepath = f.name

        try:
            self.app.entry_val_file.delete(0, "end")
            self.app.entry_val_file.insert(0, filepath)
            self.app._run_validation()

            self.assertIsNotNone(self.app.last_validation_report)
            self.assertTrue(self.app.last_validation_report.is_valid)
            self.assertEqual(self.app.last_validation_report.register_count, 1)
        finally:
            try:
                os.unlink(filepath)
            except OSError:
                pass

    def test_run_validation_invalid_file(self):
        """Tests validator tab execution on an invalid file."""
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;INVALID;;Active Power;active_power;1.0;0.0;W;4\n")
            filepath = f.name

        try:
            self.app.entry_val_file.delete(0, "end")
            self.app.entry_val_file.insert(0, filepath)
            self.app._run_validation()

            self.assertIsNotNone(self.app.last_validation_report)
            self.assertFalse(self.app.last_validation_report.is_valid)
            self.assertGreater(len(self.app.tree_issues.get_children()), 0)
        finally:
            try:
                os.unlink(filepath)
            except OSError:
                pass

    def test_export_validation_report(self):
        """Tests exporting validation report to a text file."""
        self.app.last_validation_report = ValidationReport(
            is_valid=False,
            register_count=2,
            issues=[
                ValidationIssue(
                    line=3,
                    severity="ERROR",
                    code="DUPLICATE_TAG",
                    field="Tag",
                    message="Duplicate tag",
                )
            ],
            stats={"errors": 1, "warnings": 0, "types": {}, "registers": 2},
        )
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            out_path = f.name

        try:
            with patch.object(gui_module.filedialog, "asksaveasfilename", return_value=out_path):
                with patch.object(gui_module.messagebox, "showinfo") as mock_info:
                    self.app._export_validation_report()
                    mock_info.assert_called_once()
            self.assertTrue(os.path.exists(out_path))
            with open(out_path, encoding="utf-8") as f:
                content = f.read()
            self.assertIn("RAPPORT D'AUDIT ET VALIDATION", content)
            self.assertIn("DUPLICATE_TAG", content)
        finally:
            try:
                os.unlink(out_path)
            except OSError:
                pass


class TestMainCLIGUIDispatch(unittest.TestCase):
    @patch("DefFileGenerator.gui.main")
    def test_main_gui_subcommand(self, mock_gui_main):
        """Tests that `deffilegen gui` invokes gui.main()."""
        import DefFileGenerator.main as cli_main

        with self.assertRaises(SystemExit) as cm:
            cli_main._run_cli(["gui"])
        self.assertEqual(cm.exception.code, 0)
        mock_gui_main.assert_called_once()

    @patch("DefFileGenerator.gui.main")
    def test_main_gui_flag(self, mock_gui_main):
        """Tests that `deffilegen --gui` invokes gui.main()."""
        import DefFileGenerator.main as cli_main

        with self.assertRaises(SystemExit) as cm:
            cli_main._run_cli(["--gui"])
        self.assertEqual(cm.exception.code, 0)
        mock_gui_main.assert_called_once()


class TestDefFileGenAppFallbackTkinter(unittest.TestCase):
    """Verifies that DefFileGenApp works completely when customtkinter is not available."""

    @classmethod
    def setUpClass(cls):
        try:
            with patch.object(gui_module, "HAS_CUSTOMTKINTER", False):
                cls.root = gui_module.tk.Tk()
                cls.root.withdraw()
                cls.app = DefFileGenApp(cls.root)
        except Exception as e:
            raise unittest.SkipTest(f"Tkinter environment not available: {e}") from e


    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def test_fallback_widgets_instantiation(self):
        """Verifies that all widgets, entries, and action buttons exist in fallback mode."""
        self.assertIsNotNone(self.app.entry_input_file)
        self.assertIsNotNone(self.app.entry_mfg)
        self.assertIsNotNone(self.app.entry_model)
        self.assertIsNotNone(self.app.opt_protocol)
        self.assertIsNotNone(self.app.opt_category)
        self.assertIsNotNone(self.app.entry_offset)
        self.assertIsNotNone(self.app.entry_output_file)
        self.assertIsNotNone(self.app.progress_bar)
        self.assertIsNotNone(self.app.lbl_results_badge)
        self.assertIsNotNone(self.app.btn_open_file)
        self.assertIsNotNone(self.app.btn_open_folder)
        self.assertIsNotNone(self.app.btn_copy_csv)
        self.assertIsNotNone(self.app.tree_preview)

    def test_fallback_on_conversion_success(self):
        """Verifies preview treeview population and action button activation in fallback mode."""
        mock_rows = [
            {
                "RegisterType": "Holding",
                "Address": "40001",
                "Type": "U16",
                "Name": "Solar Voltage",
                "Tag": "solar_volt",
                "Factor": "0.1",
                "Offset": "0",
                "Unit": "V",
                "Action": "4",
            }
        ]
        mock_report = ValidationReport(
            is_valid=True,
            register_count=1,
            issues=[],
            stats={"errors": 0, "warnings": 0, "types": {"3": 1}, "registers": 1},
        )
        with patch.object(gui_module.messagebox, "showinfo"):
            self.app._on_conversion_success("test_out_fallback.csv", mock_rows, mock_report)

        self.assertEqual(self.app.last_generated_file, "test_out_fallback.csv")
        self.assertEqual(len(self.app.tree_preview.get_children()), 1)
        self.assertEqual(str(self.app.btn_open_file["state"]), "normal")
        self.assertEqual(str(self.app.btn_open_folder["state"]), "normal")
        self.assertEqual(str(self.app.btn_copy_csv["state"]), "normal")


if __name__ == "__main__":
    unittest.main()

