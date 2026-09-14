#!/usr/bin/env python3
"""
Windows 11 Native Desktop Application for WebdynSunPM DefFileGenerator.

Provides a fast, robust, and modern graphical user interface designed
specifically for Windows 11 with dark/light theme support, high-DPI scaling,
multi-threaded non-blocking extraction, live register preview, full definition
validation, template generation, and real-time logging.
"""

from __future__ import annotations

import csv
import ctypes
import logging
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    ValidationReport,
    run_generator,
)
from DefFileGenerator.extractor import Extractor, peek_generator

# Optional CustomTkinter import
HAS_CUSTOMTKINTER = False
try:
    import customtkinter as ctk  # type: ignore[import-untyped, import-not-found]

    HAS_CUSTOMTKINTER = True
except ImportError:
    pass

# Enable Windows High-DPI awareness before creating UI instances
if sys.platform == "win32":
    try:
        # Per-monitor DPI awareness v2 (Windows 10 Creators Update and Windows 11)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

logger = logging.getLogger("DefFileGenerator.gui")

# Equipment categories and pre-configured templates
EQUIPMENT_TEMPLATES: dict[str, dict[str, Any]] = {
    "Inverter": {
        "description": "Solar PV Inverter (Active Power, Voltage, Frequency, Energy)",
        "registers": [
            (
                "1",
                "3",
                "40001",
                "U16",
                "",
                "Grid Voltage",
                "grid_voltage",
                "0.100000",
                "0.000000",
                "V",
                "4",
            ),
            (
                "2",
                "3",
                "40002",
                "U16",
                "",
                "Grid Frequency",
                "grid_freq",
                "0.010000",
                "0.000000",
                "Hz",
                "4",
            ),
            (
                "3",
                "3",
                "40003",
                "U32_WB",
                "",
                "Active Power",
                "active_power",
                "1.000000",
                "0.000000",
                "W",
                "4",
            ),
            (
                "4",
                "3",
                "40005",
                "U32_WB",
                "",
                "Total Energy",
                "total_energy",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            ),
            (
                "5",
                "3",
                "40007",
                "U16",
                "",
                "Internal Temp",
                "inv_temp",
                "0.100000",
                "0.000000",
                "°C",
                "4",
            ),
            (
                "6",
                "3",
                "40008",
                "U16",
                "",
                "Inverter State",
                "inv_state",
                "1.000000",
                "0.000000",
                "",
                "4",
            ),
        ],
    },
    "Meter": {
        "description": "Grid / Sub-distribution Power Meter (Energy, Voltage, Current, Cos Phi)",
        "registers": [
            (
                "1",
                "3",
                "30001",
                "F32_WB",
                "",
                "Voltage L1-N",
                "volt_l1",
                "1.000000",
                "0.000000",
                "V",
                "4",
            ),
            (
                "2",
                "3",
                "30003",
                "F32_WB",
                "",
                "Voltage L2-N",
                "volt_l2",
                "1.000000",
                "0.000000",
                "V",
                "4",
            ),
            (
                "3",
                "3",
                "30005",
                "F32_WB",
                "",
                "Voltage L3-N",
                "volt_l3",
                "1.000000",
                "0.000000",
                "V",
                "4",
            ),
            (
                "4",
                "3",
                "30007",
                "F32_WB",
                "",
                "Current Total",
                "current_tot",
                "1.000000",
                "0.000000",
                "A",
                "4",
            ),
            (
                "5",
                "3",
                "30009",
                "F32_WB",
                "",
                "Total Active Power",
                "power_tot",
                "1.000000",
                "0.000000",
                "kW",
                "4",
            ),
            (
                "6",
                "3",
                "30011",
                "F32_WB",
                "",
                "Total Active Energy",
                "energy_tot",
                "1.000000",
                "0.000000",
                "kWh",
                "4",
            ),
        ],
    },
    "Sensor": {
        "description": "Pyranometer / Irradiance & Ambient Temperature Sensor",
        "registers": [
            (
                "1",
                "3",
                "100",
                "F32_WB",
                "",
                "Solar Irradiance",
                "irradiance",
                "1.000000",
                "0.000000",
                "W/m²",
                "4",
            ),
            (
                "2",
                "3",
                "102",
                "F32_WB",
                "",
                "Sensor Temperature",
                "sensor_temp",
                "1.000000",
                "0.000000",
                "°C",
                "4",
            ),
            (
                "3",
                "3",
                "104",
                "F32_WB",
                "",
                "Ambient Temperature",
                "ambient_temp",
                "1.000000",
                "0.000000",
                "°C",
                "4",
            ),
            (
                "4",
                "3",
                "106",
                "U16",
                "",
                "Sensor Status",
                "sensor_status",
                "1.000000",
                "0.000000",
                "",
                "4",
            ),
        ],
    },
    "Battery": {
        "description": "Battery Energy Storage System (BESS - SOC, SOH, Voltage, Current, Temp)",
        "registers": [
            (
                "1",
                "3",
                "50001",
                "U16",
                "",
                "State of Charge",
                "bess_soc",
                "0.100000",
                "0.000000",
                "%",
                "4",
            ),
            (
                "2",
                "3",
                "50002",
                "U16",
                "",
                "State of Health",
                "bess_soh",
                "0.100000",
                "0.000000",
                "%",
                "4",
            ),
            (
                "3",
                "3",
                "50003",
                "U16",
                "",
                "DC Battery Voltage",
                "bess_voltage",
                "0.100000",
                "0.000000",
                "V",
                "4",
            ),
            (
                "4",
                "3",
                "50004",
                "I16",
                "",
                "DC Battery Current",
                "bess_current",
                "0.100000",
                "0.000000",
                "A",
                "4",
            ),
            (
                "5",
                "3",
                "50005",
                "I16",
                "",
                "DC Battery Power",
                "bess_power",
                "1.000000",
                "0.000000",
                "kW",
                "4",
            ),
            (
                "6",
                "3",
                "50006",
                "I16",
                "",
                "Battery Max Temp",
                "bess_max_temp",
                "0.100000",
                "0.000000",
                "°C",
                "4",
            ),
        ],
    },
    "WeatherStation": {
        "description": "Complete Meteorological Station (Irradiance, Wind Speed, Wind Direction, Rain, Humidity)",
        "registers": [
            (
                "1",
                "3",
                "1",
                "U16",
                "",
                "Global Horizontal Irradiance",
                "ghi",
                "1.000000",
                "0.000000",
                "W/m²",
                "4",
            ),
            (
                "2",
                "3",
                "2",
                "U16",
                "",
                "Plane of Array Irradiance",
                "poa",
                "1.000000",
                "0.000000",
                "W/m²",
                "4",
            ),
            (
                "3",
                "3",
                "3",
                "U16",
                "",
                "Wind Speed",
                "wind_speed",
                "0.100000",
                "0.000000",
                "m/s",
                "4",
            ),
            (
                "4",
                "3",
                "4",
                "U16",
                "",
                "Wind Direction",
                "wind_dir",
                "1.000000",
                "0.000000",
                "°",
                "4",
            ),
            (
                "5",
                "3",
                "5",
                "I16",
                "",
                "Ambient Temp",
                "weather_temp",
                "0.100000",
                "0.000000",
                "°C",
                "4",
            ),
            (
                "6",
                "3",
                "6",
                "U16",
                "",
                "Relative Humidity",
                "weather_humidity",
                "0.100000",
                "0.000000",
                "%",
                "4",
            ),
        ],
    },
    "Generic": {
        "description": "Generic Modbus RTU / TCP Device Template",
        "registers": [
            (
                "1",
                "3",
                "1",
                "U16",
                "",
                "Sample Variable 1",
                "sample_var_1",
                "1.000000",
                "0.000000",
                "",
                "4",
            ),
            (
                "2",
                "3",
                "2",
                "U16",
                "",
                "Sample Variable 2",
                "sample_var_2",
                "1.000000",
                "0.000000",
                "",
                "4",
            ),
            (
                "3",
                "3",
                "3",
                "U32_WB",
                "",
                "Sample Counter 3",
                "sample_cnt_3",
                "1.000000",
                "0.000000",
                "",
                "4",
            ),
        ],
    },
}


class QueueLogHandler(logging.Handler):
    """Thread-safe logging handler routing records into a queue for UI display."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self.log_queue.put((record.levelname, msg))


class DefFileGenApp:
    """Main Desktop Application Window designed for Windows 11."""

    def __init__(self, root: tk.Tk | ctk.CTk):
        self.root = root
        self.root.title("WebdynSunPM DefFileGenerator — Windows 11")
        self.root.geometry("1100x800")
        self.root.minsize(940, 680)

        # Logging queue
        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.log_handler = QueueLogHandler(self.log_queue)
        self.log_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        )
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)

        # State storage
        self.extracted_rows: list[dict[str, Any]] = []
        self.last_generated_file: str | None = None
        self.last_validation_report: ValidationReport | None = None
        self.is_processing = False

        self._configure_styles()
        self._build_ui()
        self._start_log_consumer()

    def _configure_styles(self) -> None:
        """Sets up Windows 11 typography and treeview styles."""
        if HAS_CUSTOMTKINTER:
            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")

        style = ttk.Style()
        # Use native 'vista' or 'clam' for crisp rendering on Windows
        available = style.theme_names()
        if "vista" in available:
            style.theme_use("vista")
        elif "clam" in available:
            style.theme_use("clam")

        # Typography configuration (Segoe UI Variable is native to Windows 11)
        font_main = ("Segoe UI Variable Display", 10)
        font_head = ("Segoe UI Variable Display", 10, "bold")

        style.configure("Treeview", font=font_main, rowheight=26)
        style.configure("Treeview.Heading", font=font_head)

    def _build_ui(self) -> None:
        """Constructs the application layout."""
        # Top Header Bar
        self.header_frame = (
            ctk.CTkFrame(self.root, height=54, corner_radius=0)
            if HAS_CUSTOMTKINTER
            else tk.Frame(self.root, height=54, bg="#1E293B")
        )
        self.header_frame.pack(side="top", fill="x")

        title_text = "⚡ WebdynSunPM DefFileGenerator"
        subtitle_text = "Windows 11 Edition · Modbus to Webdyn Definition Generator"
        if HAS_CUSTOMTKINTER:
            self.title_label = ctk.CTkLabel(
                self.header_frame,
                text=f"{title_text}  |  {subtitle_text}",
                font=ctk.CTkFont(family="Segoe UI Variable Display", size=14, weight="bold"),
            )
            self.title_label.pack(side="left", padx=20, pady=12)

            self.theme_btn = ctk.CTkButton(
                self.header_frame,
                text="🌓 Theme",
                width=80,
                height=28,
                command=self._toggle_theme,
            )
            self.theme_btn.pack(side="right", padx=20, pady=12)
        else:
            self.title_label = tk.Label(
                self.header_frame,
                text=f"{title_text} - {subtitle_text}",
                font=("Segoe UI", 12, "bold"),
                fg="#FFFFFF",
                bg="#1E293B",
            )
            self.title_label.pack(side="left", padx=20, pady=14)

        # Tab Navigation View
        if HAS_CUSTOMTKINTER:
            self.tabs = ctk.CTkTabview(self.root, corner_radius=10)
            self.tabs.pack(fill="both", expand=True, padx=16, pady=(8, 16))
            self.tab_convert = self.tabs.add("⚡ Conversion & Generation")
            self.tab_validator = self.tabs.add("🛡️ Definition Validator")
            self.tab_templates = self.tabs.add("📋 Equipment Templates")
            self.tab_logs = self.tabs.add("📜 Console & Logs")
        else:
            self.tabs = ttk.Notebook(self.root)
            self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
            self.tab_convert = ttk.Frame(self.tabs)
            self.tab_validator = ttk.Frame(self.tabs)
            self.tab_templates = ttk.Frame(self.tabs)
            self.tab_logs = ttk.Frame(self.tabs)
            self.tabs.add(self.tab_convert, text="Conversion & Generation")
            self.tabs.add(self.tab_validator, text="Validateur")
            self.tabs.add(self.tab_templates, text="Templates")
            self.tabs.add(self.tab_logs, text="Logs")

        self._build_convert_tab()
        self._build_validator_tab()
        self._build_templates_tab()
        self._build_logs_tab()

    def _toggle_theme(self) -> None:
        """Toggles between Dark and Light mode dynamically."""
        if not HAS_CUSTOMTKINTER:
            return
        current = ctk.get_appearance_mode()
        new_mode = "Light" if current == "Dark" else "Dark"
        ctk.set_appearance_mode(new_mode)

    # -------------------------------------------------------------------------
    # TAB 1: CONVERT & GENERATE
    # -------------------------------------------------------------------------
    def _build_convert_tab(self) -> None:
        parent = self.tab_convert

        # Section 1: File Ingestion Frame
        file_frame = (
            ctk.CTkFrame(parent, corner_radius=8)
            if HAS_CUSTOMTKINTER
            else ttk.LabelFrame(parent, text="Source Documentation File")
        )
        file_frame.pack(fill="x", padx=12, pady=(8, 6))

        if HAS_CUSTOMTKINTER:
            lbl_file = ctk.CTkLabel(
                file_frame,
                text="Source File (PDF, Excel .xlsx, CSV, XML) :",
                font=ctk.CTkFont(weight="bold"),
            )
            lbl_file.grid(row=0, column=0, sticky="w", padx=14, pady=(10, 2))

            self.entry_input_file = ctk.CTkEntry(
                file_frame, placeholder_text="Select or drag a file...", width=620
            )
            self.entry_input_file.grid(row=1, column=0, padx=14, pady=(0, 10), sticky="ew")

            btn_browse_in = ctk.CTkButton(
                file_frame, text="📂 Browse...", width=120, command=self._browse_input_file
            )
            btn_browse_in.grid(row=1, column=1, padx=(0, 14), pady=(0, 10))

            file_frame.columnconfigure(0, weight=1)
        else:
            self.entry_input_file = ttk.Entry(file_frame, width=80)
            self.entry_input_file.pack(side="left", fill="x", expand=True, padx=10, pady=10)
            btn_browse_in = ttk.Button(
                file_frame, text="Parcourir...", command=self._browse_input_file
            )
            btn_browse_in.pack(side="right", padx=10, pady=10)

        # Section 2: Parameters Grid Frame
        params_frame = (
            ctk.CTkFrame(parent, corner_radius=8)
            if HAS_CUSTOMTKINTER
            else ttk.LabelFrame(parent, text="WebdynSunPM Parameters")
        )
        params_frame.pack(fill="x", padx=12, pady=6)

        if HAS_CUSTOMTKINTER:
            # Row 0: Manufacturer & Model
            ctk.CTkLabel(params_frame, text="Constructeur (Manufacturer) *").grid(
                row=0, column=0, sticky="w", padx=14, pady=(10, 2)
            )
            self.entry_mfg = ctk.CTkEntry(params_frame, placeholder_text="Ex: Huawei, ABB, SMA")
            self.entry_mfg.insert(0, "Huawei")
            self.entry_mfg.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))

            ctk.CTkLabel(params_frame, text="Model *").grid(
                row=0, column=1, sticky="w", padx=14, pady=(10, 2)
            )
            self.entry_model = ctk.CTkEntry(params_frame, placeholder_text="Ex: SUN2000-50KTL")
            self.entry_model.insert(0, "SUN2000")
            self.entry_model.grid(row=1, column=1, sticky="ew", padx=14, pady=(0, 8))

            # Row 1: Protocol & Category
            ctk.CTkLabel(params_frame, text="Protocole").grid(
                row=2, column=0, sticky="w", padx=14, pady=(4, 2)
            )
            self.opt_protocol = ctk.CTkOptionMenu(params_frame, values=["modbusRTU", "modbusTCP"])
            self.opt_protocol.set("modbusRTU")
            self.opt_protocol.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 8))

            ctk.CTkLabel(params_frame, text="Category").grid(
                row=2, column=1, sticky="w", padx=14, pady=(4, 2)
            )
            self.opt_category = ctk.CTkOptionMenu(
                params_frame,
                values=[
                    "Inverter",
                    "Meter",
                    "Sensor",
                    "Battery",
                    "WeatherStation",
                    "Tracker",
                    "Other",
                ],
            )
            self.opt_category.set("Inverter")
            self.opt_category.grid(row=3, column=1, sticky="ew", padx=14, pady=(0, 8))

            # Row 2: Address Offset & Output File
            ctk.CTkLabel(params_frame, text="Address Offset").grid(
                row=4, column=0, sticky="w", padx=14, pady=(4, 2)
            )
            self.entry_offset = ctk.CTkEntry(params_frame, placeholder_text="0")
            self.entry_offset.insert(0, "0")
            self.entry_offset.grid(row=5, column=0, sticky="ew", padx=14, pady=(0, 12))

            ctk.CTkLabel(params_frame, text="Output File (.csv) :").grid(
                row=4, column=1, sticky="w", padx=14, pady=(4, 2)
            )
            out_box = ctk.CTkFrame(params_frame, fg_color="transparent")
            out_box.grid(row=5, column=1, sticky="ew", padx=14, pady=(0, 12))
            self.entry_output_file = ctk.CTkEntry(
                out_box, placeholder_text="Auto-generated or custom"
            )
            self.entry_output_file.pack(side="left", fill="x", expand=True, padx=(0, 8))
            btn_browse_out = ctk.CTkButton(
                out_box, text="Enregistrer sous...", width=110, command=self._browse_output_file
            )
            btn_browse_out.pack(side="right")

            params_frame.columnconfigure(0, weight=1)
            params_frame.columnconfigure(1, weight=1)
        else:
            # Row 0: Manufacturer & Model
            ttk.Label(params_frame, text="Constructeur (Manufacturer) *").grid(
                row=0, column=0, sticky="w", padx=10, pady=(8, 2)
            )
            self.entry_mfg = ttk.Entry(params_frame)
            self.entry_mfg.insert(0, "Huawei")
            self.entry_mfg.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))

            ttk.Label(params_frame, text="Model *").grid(
                row=0, column=1, sticky="w", padx=10, pady=(8, 2)
            )
            self.entry_model = ttk.Entry(params_frame)
            self.entry_model.insert(0, "SUN2000")
            self.entry_model.grid(row=1, column=1, sticky="ew", padx=10, pady=(0, 6))

            # Row 1: Protocol & Category
            ttk.Label(params_frame, text="Protocole").grid(
                row=2, column=0, sticky="w", padx=10, pady=(4, 2)
            )
            self.opt_protocol = ttk.Combobox(
                params_frame, values=["modbusRTU", "modbusTCP"], state="readonly"
            )
            self.opt_protocol.set("modbusRTU")
            self.opt_protocol.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 6))

            ttk.Label(params_frame, text="Category").grid(
                row=2, column=1, sticky="w", padx=10, pady=(4, 2)
            )
            self.opt_category = ttk.Combobox(
                params_frame,
                values=[
                    "Inverter",
                    "Meter",
                    "Sensor",
                    "Battery",
                    "WeatherStation",
                    "Tracker",
                    "Other",
                ],
                state="readonly",
            )
            self.opt_category.set("Inverter")
            self.opt_category.grid(row=3, column=1, sticky="ew", padx=10, pady=(0, 6))

            # Row 2: Address Offset & Output File
            ttk.Label(params_frame, text="Address Offset").grid(
                row=4, column=0, sticky="w", padx=10, pady=(4, 2)
            )
            self.entry_offset = ttk.Entry(params_frame)
            self.entry_offset.insert(0, "0")
            self.entry_offset.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 10))

            ttk.Label(params_frame, text="Output File (.csv) :").grid(
                row=4, column=1, sticky="w", padx=10, pady=(4, 2)
            )
            out_box = ttk.Frame(params_frame)
            out_box.grid(row=5, column=1, sticky="ew", padx=10, pady=(0, 10))
            self.entry_output_file = ttk.Entry(out_box)
            self.entry_output_file.pack(side="left", fill="x", expand=True, padx=(0, 6))
            btn_browse_out = ttk.Button(
                out_box, text="Enregistrer sous...", command=self._browse_output_file
            )
            btn_browse_out.pack(side="right")

            params_frame.columnconfigure(0, weight=1)
            params_frame.columnconfigure(1, weight=1)

        # Section 3: Action & Progress Bar
        action_frame = (
            ctk.CTkFrame(parent, fg_color="transparent") if HAS_CUSTOMTKINTER else ttk.Frame(parent)
        )
        action_frame.pack(fill="x", padx=12, pady=4)

        if HAS_CUSTOMTKINTER:
            self.btn_convert = ctk.CTkButton(
                action_frame,
                text="⚡ Extract & Generate WebdynSunPM Definition File",
                font=ctk.CTkFont(size=13, weight="bold"),
                height=38,
                command=self._start_conversion_thread,
            )
            self.btn_convert.pack(side="left", fill="x", expand=True, padx=(0, 10))

            self.progress_bar = ctk.CTkProgressBar(action_frame, mode="indeterminate", width=220)
            self.progress_bar.pack(side="right", padx=4)
            self.progress_bar.set(0)
        else:
            self.btn_convert = ttk.Button(
                action_frame,
                text="⚡ Extract & Generate WebdynSunPM Definition File",
                command=self._start_conversion_thread,
            )
            self.btn_convert.pack(side="left", fill="x", expand=True, padx=(0, 10))

            self.progress_bar = ttk.Progressbar(action_frame, mode="indeterminate", length=220)
            self.progress_bar.pack(side="right", padx=4)

        # Section 4: Results Preview & Quick Action Bar
        results_header = (
            ctk.CTkFrame(parent, fg_color="transparent") if HAS_CUSTOMTKINTER else ttk.Frame(parent)
        )
        results_header.pack(fill="x", padx=14, pady=(6, 2))

        if HAS_CUSTOMTKINTER:
            self.lbl_results_badge = ctk.CTkLabel(
                results_header,
                text="No registers extracted yet",
                font=ctk.CTkFont(weight="bold"),
            )
            self.lbl_results_badge.pack(side="left")

            self.btn_open_file = ctk.CTkButton(
                results_header,
                text="📄 Open CSV File",
                width=140,
                state="disabled",
                command=self._open_last_generated_file,
            )
            self.btn_open_file.pack(side="right", padx=(6, 0))

            self.btn_open_folder = ctk.CTkButton(
                results_header,
                text="📁 Ouvrir Dossier",
                width=120,
                state="disabled",
                command=self._open_output_folder,
            )
            self.btn_open_folder.pack(side="right", padx=(6, 0))

            self.btn_copy_csv = ctk.CTkButton(
                results_header,
                text="📋 Copy CSV",
                width=100,
                state="disabled",
                command=self._copy_csv_to_clipboard,
            )
            self.btn_copy_csv.pack(side="right", padx=(6, 0))
        else:
            self.lbl_results_badge = ttk.Label(
                results_header,
                text="No registers extracted yet",
                font=("Segoe UI", 9, "bold"),
            )
            self.lbl_results_badge.pack(side="left")

            self.btn_open_file = ttk.Button(
                results_header,
                text="📄 Open CSV File",
                state="disabled",
                command=self._open_last_generated_file,
            )
            self.btn_open_file.pack(side="right", padx=(6, 0))

            self.btn_open_folder = ttk.Button(
                results_header,
                text="📁 Ouvrir Dossier",
                state="disabled",
                command=self._open_output_folder,
            )
            self.btn_open_folder.pack(side="right", padx=(6, 0))

            self.btn_copy_csv = ttk.Button(
                results_header,
                text="📋 Copy CSV",
                state="disabled",
                command=self._copy_csv_to_clipboard,
            )
            self.btn_copy_csv.pack(side="right", padx=(6, 0))

        # Preview Treeview Table
        tree_container = ttk.Frame(parent)
        tree_container.pack(fill="both", expand=True, padx=12, pady=(4, 8))

        columns = (
            "index",
            "info1",
            "address",
            "type",
            "name",
            "tag",
            "coefa",
            "coefb",
            "unit",
            "action",
        )
        self.tree_preview = ttk.Treeview(
            tree_container, columns=columns, show="headings", selectmode="browse"
        )

        headers = [
            ("#", 45),
            ("RegType", 75),
            ("Adresse", 90),
            ("Type", 80),
            ("Parameter name", 260),
            ("Tag", 140),
            ("CoefA", 75),
            ("CoefB", 65),
            ("Unit", 65),
            ("Access", 60),
        ]
        for col, (name, width) in zip(columns, headers, strict=False):
            self.tree_preview.heading(col, text=name)
            self.tree_preview.column(col, width=width, anchor="w" if col == "name" else "center")

        scrollbar_y = ttk.Scrollbar(
            tree_container, orient="vertical", command=self.tree_preview.yview
        )
        scrollbar_x = ttk.Scrollbar(
            tree_container, orient="horizontal", command=self.tree_preview.xview
        )
        self.tree_preview.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree_preview.pack(fill="both", expand=True)

    def _browse_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Modbus documentation",
            filetypes=[
                ("All supported formats", "*.pdf;*.xlsx;*.xlsm;*.csv;*.xml"),
                ("PDF Documents", "*.pdf"),
                ("Excel Workbooks", "*.xlsx;*.xlsm"),
                ("CSV Files", "*.csv"),
                ("XML Files", "*.xml"),
                ("All files", "*.*"),
            ],
        )
        if file_path:
            self.entry_input_file.delete(0, tk.END)
            self.entry_input_file.insert(0, file_path)

            # Auto-infer manufacturer and model from filename
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            parts = re.split(r"[_\-\s]+", base_name)
            if parts:
                suggested_mfg = parts[0].capitalize()
                self.entry_mfg.delete(0, tk.END)
                self.entry_mfg.insert(0, suggested_mfg)
                if len(parts) > 1:
                    suggested_model = "_".join(parts[1:])
                    self.entry_model.delete(0, tk.END)
                    self.entry_model.insert(0, suggested_model)

            # Suggest default output path in same directory
            dir_name = os.path.dirname(file_path)
            clean_mfg = re.sub(r"[^a-zA-Z0-9]", "_", self.entry_mfg.get()).lower()
            clean_model = re.sub(r"[^a-zA-Z0-9]", "_", self.entry_model.get()).lower()
            suggested_out = os.path.join(dir_name, f"{clean_mfg}_{clean_model}_definition.csv")
            self.entry_output_file.delete(0, tk.END)
            self.entry_output_file.insert(0, suggested_out)

    def _browse_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Save WebdynSunPM definition",
            defaultextension=".csv",
            filetypes=[
                ("WebdynSunPM CSV Definition File", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if file_path:
            self.entry_output_file.delete(0, tk.END)
            self.entry_output_file.insert(0, file_path)

    def _start_conversion_thread(self) -> None:
        """Launches conversion in a background thread to prevent UI freezing."""
        if self.is_processing:
            return

        input_path = self.entry_input_file.get().strip()
        if not input_path:
            messagebox.showwarning("Missing file", "Please select a source file.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"The source file does not exist:\n{input_path}")
            return

        mfg = self.entry_mfg.get().strip() or "Manufacturer"
        model = self.entry_model.get().strip() or "Model"
        protocol = self.opt_protocol.get() if hasattr(self, "opt_protocol") else "modbusRTU"
        category = self.opt_category.get() if hasattr(self, "opt_category") else "Inverter"

        try:
            offset = int(self.entry_offset.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Invalid offset", "The address offset must be an integer.")
            return

        output_path = self.entry_output_file.get().strip()
        if not output_path:
            clean_mfg = re.sub(r"[^a-zA-Z0-9]", "_", mfg).lower()
            clean_model = re.sub(r"[^a-zA-Z0-9]", "_", model).lower()
            dir_name = os.path.dirname(input_path) or os.getcwd()
            output_path = os.path.join(dir_name, f"{clean_mfg}_{clean_model}_definition.csv")
            self.entry_output_file.delete(0, tk.END)
            self.entry_output_file.insert(0, output_path)

        # Update UI state to processing
        self.is_processing = True
        self.btn_convert.configure(state="disabled")
        if hasattr(self, "progress_bar"):
            self.progress_bar.start()
        self.lbl_results_badge.configure(
            text="Traitement en cours... Extraction des tables Modbus..."
        )

        # Clear existing preview rows
        for item in self.tree_preview.get_children():
            self.tree_preview.delete(item)

        thread = threading.Thread(
            target=self._run_conversion_worker,
            args=(input_path, output_path, mfg, model, protocol, category, offset),
            daemon=True,
        )
        thread.start()

    def _run_conversion_worker(
        self,
        input_path: str,
        output_path: str,
        mfg: str,
        model: str,
        protocol: str,
        category: str,
        offset: int,
    ) -> None:
        """Worker executing in background thread."""
        ext = os.path.splitext(input_path)[1].lower()
        extractor = Extractor()
        raw_data: Any = iter([])

        try:
            logger.info("Starting extraction from: %s", input_path)
            if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
                raw_data = extractor.extract_from_excel(input_path)
            elif ext == ".pdf":
                raw_data = extractor.extract_from_pdf(input_path)
            elif ext == ".csv":
                raw_data = extractor.extract_from_csv(input_path)
            elif ext == ".xml":
                raw_data = extractor.extract_from_xml(input_path)
            else:
                raise ValueError(f"Unsupported format: {ext}")

            has_data, raw_data_peeked = peek_generator(raw_data)
            if not has_data:
                raise RuntimeError(
                    "No usable data table detected in this document.\n"
                    "Ensure the document contains textual tables and not image scans."
                )

            mapped_gen = extractor.map_and_clean(raw_data_peeked, offset)
            has_regs, mapped_peeked = peek_generator(mapped_gen)
            if not has_regs:
                raise RuntimeError(
                    "No register could be mapped to standard Modbus fields.\n\n"
                    "Explications possibles :\n"
                    "• The document is a mechanical installation manual or product brochure without Modbus registers.\n"
                    "• The document does not contain usable register columns (addresses, types, names)."
                )

            full_mapped = list(mapped_peeked)

            config = GeneratorConfig(
                input_file=input_path,
                output=output_path,
                manufacturer=mfg,
                model=model,
                protocol=protocol,
                category=category,
                address_offset=0,  # Already applied in map_and_clean
            )

            run_generator(config, input_data=full_mapped)

            # Validation of the generated file
            generator = Generator()
            report = generator.validate_csv_detailed(output_path, strict=True)

            self.root.after(0, self._on_conversion_success, output_path, full_mapped, report)

        except Exception as exc:
            logger.exception("Conversion failed")
            self.root.after(0, self._on_conversion_error, str(exc))

    def _on_conversion_success(
        self, output_path: str, mapped_rows: list[dict[str, Any]], report: ValidationReport
    ) -> None:
        try:
            self.is_processing = False
            self.btn_convert.configure(state="normal")
            if hasattr(self, "progress_bar"):
                self.progress_bar.stop()
                if hasattr(self.progress_bar, "set"):
                    self.progress_bar.set(1.0)

            self.last_generated_file = output_path
            self.btn_open_file.configure(state="normal")
            self.btn_open_folder.configure(state="normal")
            self.btn_copy_csv.configure(state="normal")

            count = len(mapped_rows)
            status_text = (
                f"✅ Success: {count} registers extracted · 100% Valid WebdynSunPM File"
                if report.is_valid
                else f"⚠️ Partial success: {count} registers extracted · {len(report.issues)} warning(s)"
            )
            self.lbl_results_badge.configure(text=status_text)

            # Clear and populate treeview
            for item in self.tree_preview.get_children():
                self.tree_preview.delete(item)

            for idx, row in enumerate(mapped_rows[:1000], start=1):
                self.tree_preview.insert(
                    "",
                    "end",
                    values=(
                        idx,
                        row.get("RegisterType", "Holding"),
                        row.get("Address", ""),
                        row.get("Type", "U16"),
                        row.get("Name", ""),
                        row.get("Tag", ""),
                        row.get("Factor", "1"),
                        row.get("Offset", "0"),
                        row.get("Unit", ""),
                        row.get("Action", "4"),
                    ),
                )

            messagebox.showinfo(
                "Generation Successful",
                f"The WebdynSunPM definition file was successfully generated:\n\n{output_path}\n\n"
                f"Extracted registers: {count}\nValidation status: {'VALID' if report.is_valid else 'WARNING'}",
            )
        except Exception as exc:
            logger.exception("Error updating display")
            messagebox.showerror(
                "Display Error",
                f"Error displaying results in interface:\n{exc}",
            )

    def _on_conversion_error(self, error_msg: str) -> None:
        self.is_processing = False
        self.btn_convert.configure(state="normal")
        if hasattr(self, "progress_bar"):
            self.progress_bar.stop()
            if hasattr(self.progress_bar, "set"):
                self.progress_bar.set(0)
        self.lbl_results_badge.configure(text=f"❌ Error : {error_msg}")
        messagebox.showerror("Extraction error", f"Unable to extract registers:\n\n{error_msg}")

    def _open_last_generated_file(self) -> None:
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            if sys.platform == "win32":
                os.startfile(self.last_generated_file)  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", self.last_generated_file], check=False)

    def _open_output_folder(self) -> None:
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            folder = os.path.dirname(os.path.abspath(self.last_generated_file))
            if sys.platform == "win32":
                os.startfile(folder)  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", folder], check=False)

    def _copy_csv_to_clipboard(self) -> None:
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            with open(self.last_generated_file, encoding="utf-8-sig") as f:
                content = f.read()
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            messagebox.showinfo(
                "Copied", "The definition CSV content has been copied to the clipboard."
            )

    # -------------------------------------------------------------------------
    # TAB 2: DEFINITION VALIDATOR
    # -------------------------------------------------------------------------
    def _build_validator_tab(self) -> None:
        parent = self.tab_validator

        # File selection
        val_frame = (
            ctk.CTkFrame(parent, corner_radius=8)
            if HAS_CUSTOMTKINTER
            else ttk.LabelFrame(parent, text="Definition File to Validate")
        )
        val_frame.pack(fill="x", padx=12, pady=10)

        if HAS_CUSTOMTKINTER:
            ctk.CTkLabel(
                val_frame,
                text="Select a WebdynSunPM definition file (.csv):",
                font=ctk.CTkFont(weight="bold"),
            ).pack(anchor="w", padx=14, pady=(10, 2))

            box = ctk.CTkFrame(val_frame, fg_color="transparent")
            box.pack(fill="x", padx=14, pady=(0, 10))

            self.entry_val_file = ctk.CTkEntry(
                box, placeholder_text="Path to the .csv file...", width=600
            )
            self.entry_val_file.pack(side="left", fill="x", expand=True, padx=(0, 10))

            btn_browse_val = ctk.CTkButton(
                box, text="📂 Browse...", width=120, command=self._browse_validator_file
            )
            btn_browse_val.pack(side="left", padx=(0, 8))

            btn_run_val = ctk.CTkButton(
                box,
                text="🛡️ Valider Maintenant",
                width=160,
                font=ctk.CTkFont(weight="bold"),
                command=self._run_validation,
            )
            btn_run_val.pack(side="left")
        else:
            ttk.Label(
                val_frame,
                text="Select a WebdynSunPM definition file (.csv):",
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w", padx=10, pady=(8, 2))

            box = ttk.Frame(val_frame)
            box.pack(fill="x", padx=10, pady=(0, 10))

            self.entry_val_file = ttk.Entry(box, width=70)
            self.entry_val_file.pack(side="left", fill="x", expand=True, padx=(0, 8))

            btn_browse_val = ttk.Button(
                box, text="📂 Browse...", command=self._browse_validator_file
            )
            btn_browse_val.pack(side="left", padx=(0, 6))

            btn_run_val = ttk.Button(
                box,
                text="🛡️ Valider Maintenant",
                command=self._run_validation,
            )
            btn_run_val.pack(side="left")

        # Summary Banner
        self.banner_frame = (
            ctk.CTkFrame(parent, corner_radius=8) if HAS_CUSTOMTKINTER else ttk.Frame(parent)
        )
        self.banner_frame.pack(fill="x", padx=12, pady=(0, 8))

        if HAS_CUSTOMTKINTER:
            self.lbl_val_status = ctk.CTkLabel(
                self.banner_frame,
                text="No file validated. Select a WebdynSunPM definition above.",
                font=ctk.CTkFont(size=13, weight="bold"),
            )
            self.lbl_val_status.pack(side="left", padx=16, pady=10)

            self.btn_export_val = ctk.CTkButton(
                self.banner_frame,
                text="💾 Exporter Rapport",
                width=140,
                state="disabled",
                command=self._export_validation_report,
            )
            self.btn_export_val.pack(side="right", padx=14, pady=10)
        else:
            self.lbl_val_status = ttk.Label(
                self.banner_frame,
                text="No file validated. Select a WebdynSunPM definition above.",
                font=("Segoe UI", 10, "bold"),
            )
            self.lbl_val_status.pack(side="left", padx=12, pady=8)

            self.btn_export_val = ttk.Button(
                self.banner_frame,
                text="💾 Exporter Rapport",
                state="disabled",
                command=self._export_validation_report,
            )
            self.btn_export_val.pack(side="right", padx=12, pady=8)

        # Issues Table
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        columns = ("line", "severity", "code", "field", "message")
        self.tree_issues = ttk.Treeview(
            table_frame, columns=columns, show="headings", selectmode="browse"
        )

        headers = [
            ("Ligne", 65),
            ("Severity", 90),
            ("Code", 130),
            ("Champ", 100),
            ("Message / Error Description", 550),
        ]
        for col, (name, width) in zip(columns, headers, strict=False):
            self.tree_issues.heading(col, text=name)
            self.tree_issues.column(col, width=width, anchor="w" if col == "message" else "center")

        scrollbar_val = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree_issues.yview
        )
        self.tree_issues.configure(yscrollcommand=scrollbar_val.set)
        scrollbar_val.pack(side="right", fill="y")
        self.tree_issues.pack(fill="both", expand=True)

    def _browse_validator_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select a WebdynSunPM definition",
            filetypes=[("CSV Definition File", "*.csv"), ("All files", "*.*")],
        )
        if file_path:
            self.entry_val_file.delete(0, tk.END)
            self.entry_val_file.insert(0, file_path)
            self._run_validation()

    def _run_validation(self) -> None:
        filepath = self.entry_val_file.get().strip()
        if not filepath or not os.path.exists(filepath):
            messagebox.showwarning("File not found", "Please select an existing CSV file.")
            return

        generator = Generator()
        report = generator.validate_csv_detailed(filepath, strict=True)
        self.last_validation_report = report

        # Clear existing table
        for item in self.tree_issues.get_children():
            self.tree_issues.delete(item)

        # Update status banner
        if report.is_valid:
            status_text = (
                f"✅ 100% COMPLIANT DEFINITION · {report.register_count} valid registers · 0 errors"
            )
            if HAS_CUSTOMTKINTER:
                self.lbl_val_status.configure(
                    text=status_text,
                    text_color="#10B981",  # Emerald green
                )
            else:
                self.lbl_val_status.configure(text=status_text)
        else:
            errs = report.stats.get("errors", 0)
            warns = report.stats.get("warnings", 0)
            status_text = f"❌ NON-COMPLIANT DEFINITION · {report.register_count} registers · {errs} Critical Error(s) · {warns} Warning(s)"
            if HAS_CUSTOMTKINTER:
                self.lbl_val_status.configure(
                    text=status_text,
                    text_color="#EF4444",  # Crimson red
                )
            else:
                self.lbl_val_status.configure(text=status_text)

        self.btn_export_val.configure(state="normal")

        # Insert issues into tree
        if not report.issues:
            self.tree_issues.insert(
                "",
                "end",
                values=(
                    "-",
                    "INFO",
                    "CLEAN",
                    "-",
                    "No problems detected. The file complies with the Webdyn specification.",
                ),
            )
        else:
            for issue in report.issues:
                self.tree_issues.insert(
                    "",
                    "end",
                    values=(
                        issue.line if issue.line > 0 else "-",
                        issue.severity,
                        issue.code,
                        issue.field,
                        issue.message,
                    ),
                )

    def _export_validation_report(self) -> None:
        if not self.last_validation_report:
            return
        out_path = filedialog.asksaveasfilename(
            title="Export validation report",
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt"), ("JSON File", "*.json")],
        )
        if not out_path:
            return

        report = self.last_validation_report
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("=" * 70 + "\n")
            f.write("WEBDYNSUNPM AUDIT AND VALIDATION REPORT\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Status : {'VALID' if report.is_valid else 'NON-COMPLIANT'}\n")
            f.write(f"Total registers: {report.register_count}\n")
            f.write(f"Errors : {report.stats.get('errors', 0)}\n")
            f.write(f"Warnings: {report.stats.get('warnings', 0)}\n\n")
            f.write("-" * 70 + "\n")
            f.write("DETAILS OF DETECTED ANOMALIES:\n")
            f.write("-" * 70 + "\n")
            for issue in report.issues:
                f.write(
                    f"[{issue.severity}] Ligne {issue.line} | Code: {issue.code} | Champ: {issue.field}\n"
                    f"       --> {issue.message}\n"
                )
        messagebox.showinfo("Export successful", f"Validation report saved as:\n{out_path}")

    # -------------------------------------------------------------------------
    # TAB 3: EQUIPMENT TEMPLATES
    # -------------------------------------------------------------------------
    def _build_templates_tab(self) -> None:
        parent = self.tab_templates

        card = (
            ctk.CTkFrame(parent, corner_radius=8)
            if HAS_CUSTOMTKINTER
            else ttk.LabelFrame(parent, text="WebdynSunPM Template Generator")
        )
        card.pack(fill="x", padx=14, pady=12)

        if HAS_CUSTOMTKINTER:
            ctk.CTkLabel(
                card,
                text="Pre-structured Solar & Energy Equipment Templates",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(anchor="w", padx=16, pady=(12, 4))

            ctk.CTkLabel(
                card,
                text="Instantly generate compliant, ready-to-use WebdynSunPM definitions :",
            ).pack(anchor="w", padx=16, pady=(0, 10))

            form_grid = ctk.CTkFrame(card, fg_color="transparent")
            form_grid.pack(fill="x", padx=16, pady=(0, 12))

            ctk.CTkLabel(form_grid, text="Equipment Type :", font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=0, sticky="w", padx=(0, 10), pady=6
            )
            self.opt_tmpl_category = ctk.CTkOptionMenu(
                form_grid,
                values=list(EQUIPMENT_TEMPLATES.keys()),
                command=self._on_template_selected,
            )
            self.opt_tmpl_category.set("Inverter")
            self.opt_tmpl_category.grid(row=0, column=1, sticky="w", pady=6)

            self.lbl_tmpl_desc = ctk.CTkLabel(
                form_grid,
                text=EQUIPMENT_TEMPLATES["Inverter"]["description"],
                text_color="#94A3B8",
            )
            self.lbl_tmpl_desc.grid(row=0, column=2, sticky="w", padx=(14, 0), pady=6)

            btn_gen_tmpl = ctk.CTkButton(
                card,
                text="📋 Generate and Save CSV Template",
                font=ctk.CTkFont(weight="bold"),
                height=36,
                command=self._generate_template_action,
            )
            btn_gen_tmpl.pack(anchor="w", padx=16, pady=(4, 14))
        else:
            ttk.Label(
                card,
                text="Pre-structured Solar & Energy Equipment Templates",
                font=("Segoe UI", 11, "bold"),
            ).pack(anchor="w", padx=12, pady=(10, 4))

            ttk.Label(
                card,
                text="Instantly generate compliant, ready-to-use WebdynSunPM definitions :",
            ).pack(anchor="w", padx=12, pady=(0, 8))

            form_grid = ttk.Frame(card)
            form_grid.pack(fill="x", padx=12, pady=(0, 10))

            ttk.Label(form_grid, text="Equipment Type :", font=("Segoe UI", 9, "bold")).grid(
                row=0, column=0, sticky="w", padx=(0, 8), pady=6
            )
            self.opt_tmpl_category = ttk.Combobox(
                form_grid, values=list(EQUIPMENT_TEMPLATES.keys()), state="readonly"
            )
            self.opt_tmpl_category.set("Inverter")
            self.opt_tmpl_category.grid(row=0, column=1, sticky="w", pady=6)
            self.opt_tmpl_category.bind(
                "<<ComboboxSelected>>",
                lambda _e: self._on_template_selected(self.opt_tmpl_category.get()),
            )

            self.lbl_tmpl_desc = ttk.Label(
                form_grid,
                text=EQUIPMENT_TEMPLATES["Inverter"]["description"],
            )
            self.lbl_tmpl_desc.grid(row=0, column=2, sticky="w", padx=(12, 0), pady=6)

            btn_gen_tmpl = ttk.Button(
                card,
                text="📋 Generate and Save CSV Template",
                command=self._generate_template_action,
            )
            btn_gen_tmpl.pack(anchor="w", padx=12, pady=(4, 12))

    def _on_template_selected(self, choice: str) -> None:
        if choice in EQUIPMENT_TEMPLATES:
            self.lbl_tmpl_desc.configure(text=EQUIPMENT_TEMPLATES[choice]["description"])

    def _generate_template_action(self) -> None:
        category = (
            self.opt_tmpl_category.get() if hasattr(self, "opt_tmpl_category") else "Inverter"
        )
        tmpl_data = EQUIPMENT_TEMPLATES.get(category, EQUIPMENT_TEMPLATES["Inverter"])

        save_path = filedialog.asksaveasfilename(
            title=f"Save template {category}",
            defaultextension=".csv",
            initialfile=f"modele_webdyn_{category.lower()}.csv",
            filetypes=[("WebdynSunPM CSV Definition File", "*.csv")],
        )
        if not save_path:
            return

        with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";", lineterminator="\n")
            # Header row
            header_row = ["modbusRTU", category, f"Example_{category}", "Model_1", ""] + [""] * 6
            writer.writerow(header_row)
            # Register rows
            for r in tmpl_data["registers"]:
                writer.writerow(r)

        messagebox.showinfo(
            "Generated Template",
            f"The template for '{category}' was created successfully:\n\n{save_path}\n\n"
            f"{len(tmpl_data['registers'])} standard registers configured.",
        )

    # -------------------------------------------------------------------------
    # TAB 4: REAL-TIME LOGS
    # -------------------------------------------------------------------------
    def _build_logs_tab(self) -> None:
        parent = self.tab_logs

        btn_bar = (
            ctk.CTkFrame(parent, fg_color="transparent") if HAS_CUSTOMTKINTER else ttk.Frame(parent)
        )
        btn_bar.pack(fill="x", padx=12, pady=6)

        if HAS_CUSTOMTKINTER:
            btn_clear = ctk.CTkButton(
                btn_bar, text="🗑️ Clear logs", width=120, command=self._clear_logs
            )
            btn_clear.pack(side="left", padx=(0, 8))

            btn_copy = ctk.CTkButton(
                btn_bar, text="📋 Copy logs", width=120, command=self._copy_logs
            )
            btn_copy.pack(side="left")
        else:
            btn_clear = ttk.Button(btn_bar, text="🗑️ Clear logs", command=self._clear_logs)
            btn_clear.pack(side="left", padx=(0, 8))

            btn_copy = ttk.Button(btn_bar, text="📋 Copy logs", command=self._copy_logs)
            btn_copy.pack(side="left")

        # Scrolled text widget for console
        if HAS_CUSTOMTKINTER:
            self.txt_logs = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=11))
            self.txt_logs.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        else:
            self.txt_logs = tk.Text(parent, font=("Consolas", 10), bg="#0F172A", fg="#F8FAFC")
            self.txt_logs.pack(fill="both", expand=True, padx=10, pady=10)

    def _start_log_consumer(self) -> None:
        """Consumes queue messages on the main UI thread via after()."""
        try:
            while True:
                level, msg = self.log_queue.get_nowait()
                self._append_log_message(level, msg)
        except queue.Empty:
            pass
        self.root.after(100, self._start_log_consumer)

    def _append_log_message(self, level: str, msg: str) -> None:
        prefix = "🔵 " if level == "INFO" else ("🟠 " if level == "WARNING" else "🔴 ")
        formatted = f"{prefix}{msg}\n"
        if HAS_CUSTOMTKINTER:
            self.txt_logs.insert("end", formatted)
            self.txt_logs.see("end")
        else:
            self.txt_logs.insert(tk.END, formatted)
            self.txt_logs.see(tk.END)

    def _clear_logs(self) -> None:
        if HAS_CUSTOMTKINTER:
            self.txt_logs.delete("1.0", "end")
        else:
            self.txt_logs.delete("1.0", tk.END)

    def _copy_logs(self) -> None:
        if HAS_CUSTOMTKINTER:
            logs = self.txt_logs.get("1.0", "end")
        else:
            logs = self.txt_logs.get("1.0", tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(logs)
        messagebox.showinfo("Copied", "All logs have been copied to the clipboard.")


def main() -> None:
    """Entry point for Windows 11 Desktop Application."""
    if HAS_CUSTOMTKINTER:
        root = ctk.CTk()
    else:
        root = tk.Tk()

    _app = DefFileGenApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
