# WebdynSunPM DefFileGenerator & Documentation Parser

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.10+" />
  <img src="https://img.shields.io/badge/Platform-Windows%2011%20%7C%20Linux%20%7C%20macOS-0078D4?style=for-the-badge&logo=windows&logoColor=white" alt="Platforms" />
  <img src="https://img.shields.io/badge/Version-0.3.0-blue?style=for-the-badge" alt="Version 0.3.0" />
  <img src="https://img.shields.io/badge/Tests-1082%20passed-brightgreen?style=for-the-badge&logo=pytest&logoColor=white" alt="Tests: 1082 passed" />
  <img src="https://img.shields.io/badge/Coverage-90%25-brightgreen?style=for-the-badge" alt="Coverage: 90%" />
  <img src="https://img.shields.io/badge/Type%20Check-Mypy%20Strict-blue?style=for-the-badge" alt="Mypy Strict" />
  <img src="https://img.shields.io/badge/Code%20Style-Ruff-black?style=for-the-badge" alt="Ruff" />
  <img src="https://img.shields.io/badge/Security-Bandit%20Pass-success?style=for-the-badge" alt="Bandit Security" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="MIT License" />
</p>

---

> **Universal and intelligent Modbus register table converter.**  
> Automatically extracts Modbus register maps from heterogeneous manufacturer documentation (**PDF, Excel XLSX/XLSM, CSV, XML**) and generates validated, production-ready CSV definition files for **WebdynSunPM** telemetry gateways.

---

## 📑 Table of Contents

- [🎯 What is DefFileGenerator?](#-what-is-deffilegenerator)
- [🚀 A-to-Z Installation Guide (Absolute Beginner Friendly)](#-a-to-z-installation-guide-absolute-beginner-friendly)
  - [Method 1 — Ready-to-use Standalone Executable (Recommended for Windows)](#method-1--ready-to-use-standalone-executable-recommended-for-windows)
  - [Method 2 — Express Start via `launch_gui.bat` Script](#method-2--express-start-via-launch_guibat-script)
  - [Method 3 — Complete Step-by-Step Python Installation (The Ultimate Walkthrough)](#method-3--complete-step-by-step-python-installation-the-ultimate-walkthrough)
  - [Method 4 — For macOS and Linux Users](#method-4--for-macos-and-linux-users)
- [⚡ Quick Start in 3 Minutes (First Conversion Walkthrough)](#-quick-start-in-3-minutes-first-conversion-walkthrough)
  - [Case A: Using the Windows Desktop GUI](#case-a-using-the-windows-desktop-gui)
  - [Case B: Using the Command Line Interface (CLI)](#case-b-using-the-command-line-interface-cli)
  - [Case C: Using the Local Web Browser Interface](#case-c-using-the-local-web-browser-interface)
- [🖥️ 4 Ways to Use the Software](#️-4-ways-to-use-the-software)
  - [1. Native Windows 11 Desktop GUI](#1-native-windows-11-desktop-gui)
  - [2. High-Performance CLI (`deffilegen`)](#2-high-performance-cli-deffilegen)
  - [3. Local Web Interface (FastAPI Browser App)](#3-local-web-interface-fastapi-browser-app)
  - [4. Python Programmatic API](#4-python-programmatic-api)
- [🧠 Semantic Knowledge Miner & SunPM Tagging Engine](#-semantic-knowledge-miner--sunpm-tagging-engine)
- [📊 Technical Audit & Quality Reporting Engine](#-technical-audit--quality-reporting-engine)
- [🔋 Industrial Equipment Suites (Mars Renewable & FoxESS)](#-industrial-equipment-suites-mars-renewable--foxess)
- [🏛️ Apollo-Standard Engineering Architecture & Ledgers](#️-apollo-standard-engineering-architecture--ledgers)
- [📊 System Architecture & Pipeline](#-system-architecture--pipeline)
- [📑 Supported Input Formats & Modbus Semantics](#-supported-input-formats--modbus-semantics)
- [🚨 Beginner SOS & Troubleshooting FAQ](#-beginner-sos--troubleshooting-faq)
- [🛡️ Security & Data Integrity Model](#️-security--data-integrity-model)
- [🧪 Quality Gates & Automated Tests](#-quality-gates--automated-tests)
- [📚 Full Documentation Index](#-full-documentation-index)
- [📂 Repository Structure](#-repository-structure)
- [📄 License](#-license)

---

## 🎯 What is DefFileGenerator?

To connect an industrial photovoltaic inverter (Huawei, SMA, SolarEdge, Sungrow, Fronius, ABB, etc.), energy meter, or weather station to a **WebdynSunPM** data logger, you must provide a strict semicolon-delimited CSV definition file containing every Modbus register address, data type, physical unit, and scaling factor.

Manually creating this file usually requires hours of tedious, error-prone manual work: copy-pasting tables from PDF user manuals or Excel spreadsheets, converting hexadecimal addresses to decimal numbers, calculating scale factors, and checking for overlapping memory registers.

**DefFileGenerator automates the entire process in seconds:**
1. **Reads your source document** (PDF datasheet, Excel workbook, CSV table, or XML register map).
2. **Automatically locates register tables** and understands vendor-specific column names in English, French, and German.
3. **Normalizes data types and addresses** (hex/decimal conversions, standard types like `U16`, `I32`, `F32`, individual bitfield slices).
4. **Applies Semantic SunPM Canonical Tags** via a 211,000+ register training dataset (`P_AC`, `Q_AC`, `E_TOT`, `E_DAY`, `U_AC_1`, etc.).
5. **Performs strict validation** (unique variable tags, valid Modbus address spaces, interval collision and overlap checks).
6. **Generates an official WebdynSunPM CSV file**, ready to be imported directly into your data logger gateway!

---

## 🚀 A-to-Z Installation Guide (Absolute Beginner Friendly)

Choose **the method that best matches your setup and comfort level**:

```text
┌────────────────────────────────────────────────────────────────────────┐
│                        WHICH METHOD SHOULD YOU PICK?                   │
├────────────────────────────────────────────────────────────────────────┤
│ • You want zero hassle on Windows without installing anything extra:   │
│   👉 Choose METHOD 1 (Standalone pre-compiled .exe — no Python needed) │
│                                                                        │
│ • You downloaded the project repository source folder on Windows:      │
│   👉 Choose METHOD 2 (Double-click launch_gui.bat)                     │
│                                                                        │
│ • You want to install and run the tool properly via Python / uv:       │
│   👉 Follow METHOD 3 (The complete step-by-step beginner guide)        │
│                                                                        │
│ • You are using macOS or Linux:                                        │
│   👉 Follow METHOD 4                                                   │
└────────────────────────────────────────────────────────────────────────┘
```

---

### Method 1 — Ready-to-use Standalone Executable (Recommended for Windows)

This method requires **no Python installation, no terminal commands, and zero technical background**. The application is pre-compiled as a standalone 64-bit Windows executable.

#### Step 1: Download the software
1. Navigate to the official [GitHub Releases Page](https://github.com/Cylae/DefFileGenerator/releases).
2. Locate the latest release (tagged at the top of the page).
3. Under the **Assets** section, click on the download link for:  
   `DefFileGenerator-vX.X.X-windows-x64.zip` (approximately 50 MB).
4. Save the file to your computer (e.g. in your `Downloads` folder).

#### Step 2: Extract the ZIP archive

> [!CAUTION]
> **CRITICAL BEGINNER TRAP #1:**  
> Do **NOT** double-click the executable directly from inside the `.zip` archive! Windows will open it in a temporary read-only state where auxiliary libraries cannot load, causing an immediate crash. You **MUST** extract the folder first.

1. Go to your `Downloads` folder.
2. **Right-click** on `DefFileGenerator-vX.X.X-windows-x64.zip`.
3. Select **Extract All...** (or your preferred unzip utility like 7-Zip).
4. Ensure the checkbox *"Show extracted files when complete"* is checked, then click **Extract**.
5. A new, regular uncompressed folder will open.

#### Step 3: Launch the Desktop Application
Inside the extracted folder, find:
- **`DefFileGenerator-GUI.exe`** (with the application icon).
- Double-click it to start the visual desktop interface.

#### Step 4: If Windows SmartScreen displays an alert

> [!NOTE]
> Windows Defender SmartScreen frequently shows a blue alert screen on newly downloaded open-source tools:  
> *« Windows protected your PC - Microsoft Defender SmartScreen prevented an unrecognized app from starting. »*  
> **This is completely normal** for independent open-source software that is not signed with an expensive corporate certificate.

To proceed:
1. Click the underlined text link: **"More info"** (located under the warning message).
2. A button labeled **"Run anyway"** will appear in the bottom right corner: click it.
3. The application will launch smoothly!

> [!TIP]
> If you prefer command-line usage, the exact same folder also provides **`deffilegen.exe`**. Simply open PowerShell or Command Prompt in that folder and run:  
> `.\deffilegen.exe --help`

---

### Method 2 — Express Start via `launch_gui.bat` Script

If you downloaded or cloned the project repository on a Windows computer with Python installed:

1. Open the repository root folder in Windows File Explorer.
2. Double-click the file named **`launch_gui.bat`**.
3. The script automatically checks for an active `.venv` or system Python and launches the native Windows 11 GUI without leaving an empty command prompt window open.

---

### Method 3 — Complete Step-by-Step Python Installation (The Ultimate Walkthrough)

Want to install the modifiable Python package, run the CLI, or use the local web server? Follow these clear steps.

#### Step 0: Check if Python is already installed on your PC
1. Press the **Windows Key + R** shortcut on your keyboard.
2. A small *Run* window will open in the bottom-left corner: type `cmd` and press **Enter**.
3. In the black console window, type:
   ```cmd
   python --version
   ```
4. Check the output:
   - If you see `Python 3.10.x`, `Python 3.11.x`, `Python 3.12.x`, or `Python 3.13.x`: **Great! Skip directly to Step 2.**
   - If Windows opens the Microsoft Store or prints *"'python' is not recognized"*: **Proceed to Step 1 below.**

#### Step 1: Install Python (Watch out for the #1 beginner trap!)
1. Visit the official Python download page: [python.org/downloads](https://www.python.org/downloads/).
2. Click the yellow button **Download Python** (version $\ge$ 3.10).
3. Run the downloaded installer `.exe` file.
4. ⚠️ **ATTENTION — THE MOST CRITICAL STEP:**  
   At the very bottom of the first installer screen, **you MUST check the box:**  
   ☑️ **Add python.exe to PATH**
5. Click **Install Now**.
6. When the installation finishes, click **Close**.
7. Close your command prompt and open a new one: typing `python --version` should now show your installed version!

#### Step 2: Download the project source code
Choose either option:

- **Option A (Without Git — Easiest for beginners):**
  1. Visit the repository page: [github.com/Cylae/DefFileGenerator](https://github.com/Cylae/DefFileGenerator).
  2. Click the green **`<> Code`** button near the top right.
  3. Click **`Download ZIP`**.
  4. Right-click the downloaded ZIP file > **Extract All...** to an easy-to-find location (e.g. `C:\DefFileGenerator`).

- **Option B (Using Git):**
  Open your terminal and run:
  ```powershell
  git clone https://github.com/Cylae/DefFileGenerator.git
  cd DefFileGenerator
  ```

#### Step 3: Magic trick to open terminal directly in the project folder

> [!TIP]
> **Windows Magic Shortcut to never type long directory paths:**  
> 1. Open your extracted `DefFileGenerator` folder in Windows File Explorer.  
> 2. Click on the **address bar** at the very top (where the folder path is displayed).  
> 3. Delete the text, type `powershell` (or `cmd`), and hit **Enter**.  
> 4. PowerShell will open immediately, pre-positioned in your project directory!

#### Step 4: Create and activate an isolated virtual environment

An isolated virtual environment ensures that project libraries do not conflict with anything else on your computer.

In your PowerShell window, run:

```powershell
# 1. Create the virtual environment in a folder named .venv
python -m venv .venv

# 10. Activate the virtual environment
.\.venv\Scripts\Activate.ps1
```

> [!WARNING]
> **CRITICAL BEGINNER TRAP #2 (PowerShell Execution Policy):**  
> If PowerShell prints red text stating:  
> *« File ...\Activate.ps1 cannot be loaded because running scripts is disabled on this system... »*  
> Do not worry; this is a default Windows policy. Fix it instantly by running:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
> Then re-run the activation command:
> ```powershell
> .\.venv\Scripts\Activate.ps1
> ```
> You will now see `(.venv)` displayed at the start of your command prompt. Your environment is active!

#### Step 5: Install DefFileGenerator and its dependencies

Execute the following commands:

```powershell
# Upgrade pip to the latest release
python -m pip install --upgrade pip

# Install DefFileGenerator with both CLI and Windows Desktop GUI
python -m pip install -e .

# (Optional) If you also want the browser-based Web interface:
python -m pip install -e ".[web]"
```

*Or using modern `uv` (recommended for ultra-fast setup):*
```powershell
# Bootstrap using the provided cross-platform script:
.\scripts\bootstrap.ps1
```

#### Step 6: Verify the installation!

Run:
```powershell
deffilegen --version
```
The terminal will respond:
```text
deffilegen 0.3.0
```

You are all set! 🎉

---

### Method 4 — For macOS and Linux Users

On macOS or Linux distributions (Ubuntu, Debian, Fedora, Arch):

```bash
# 1. Clone the repository
git clone https://github.com/Cylae/DefFileGenerator.git
cd DefFileGenerator

# 10. Run the bootstrap script or manually set up venv
./scripts/bootstrap.sh

# Or manual setup:
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -e ".[web]"

# 10. Verify installation
deffilegen --version
```

*(Note: The desktop GUI is optimized for Windows 11. On Linux and macOS, the CLI and the local Web browser interface are recommended).*

---

## ⚡ Quick Start in 3 Minutes (First Conversion Walkthrough)

Validate your installation immediately using the sample files provided in the repository!

### Case A: Using the Windows Desktop GUI

1. Launch the application by double-clicking **`DefFileGenerator-GUI.exe`** or **`launch_gui.bat`** (or execute `deffilegen-gui` in your activated terminal).
2. The modern Windows 11 application window opens.
3. Click **Browse** on the *Source Document* row and select the included sample file:  
   `sample_inverter_registers.xlsx`. You can also create a CSV template with `deffilegen generate --template -o template.csv`.
4. Notice that **Manufacturer** and **Model** fields auto-fill if detected from the filename.
5. Set your destination CSV file path (or keep the suggested default).
6. Click the green **Generate Definition File** button.
7. When conversion completes, inspect the logs and the preview of the generated registers.
8. Click the **Validator** tab to run instant verification on your generated file.

---

### Case B: Using the Command Line Interface (CLI)

Open your terminal in the project directory and run your first conversion:

```powershell
# Direct conversion of the bundled sample Excel register map
deffilegen run sample_inverter_registers.xlsx --manufacturer "MyBrand" --model "Inverter-5000" -o my_inverter.csv

# Immediate audit of the generated CSV file
deffilegen validate my_inverter.csv

# Generate a technical Markdown report
deffilegen report my_inverter.csv -o my_inverter_report.md
```

The CLI outputs a clear validation audit report:
```text
INFO: Validation successful: my_inverter.csv
```

---

### Case C: Using the Local Web Browser Interface

If you installed the `[web]` optional dependencies, start the local server:

```powershell
uvicorn web.app:app --host 127.0.0.1 --port 8000
```

1. Open your web browser (Chrome, Edge, Firefox, Safari).
2. Navigate to: [http://127.0.0.1:8000](http://127.0.0.1:8000).
3. Drag and drop your manufacturer document into the upload zone, configure equipment metadata, and click **Convert**.
4. Download your validated WebdynSunPM CSV file instantly!
5. To stop the web server, press **Ctrl + C** in the terminal.

---

## 🖥️ 4 Ways to Use the Software

| Interface | Best For | Key Strengths | How to Launch |
|---|---|---|---|
| **Windows Desktop GUI** | Everyone / Beginners | Visual interface, live table preview, dark/light mode, equipment templates | `DefFileGenerator-GUI.exe` or `launch_gui.bat` |
| **Command Line (CLI)** | Engineers & Automation | Fully scriptable, fast, selective PDF pages and Excel worksheet targeting | `deffilegen run ...` |
| **Local Web Interface** | Cross-platform / Browsers | Drag-and-drop web UI, responsive preview, OpenAPI `/docs` endpoints | `uvicorn web.app:app` |
| **Python Programmatic API** | Integrators & Developers | Clean functional Python API for data ingestion pipelines | `from generate_webdyn_def import generate_webdyn_definition` |

---

### 1. Native Windows 11 Desktop GUI

Built using **CustomTkinter** to provide a seamless, modern desktop experience:
- 🎨 **Adaptive Light / Dark Themes** matching Windows 11 system preferences.
- ⚡ **Background Worker Threads**: document conversion runs outside the main UI thread.
- 📋 **Pre-configured Equipment Templates**: One-click presets for Solar PV Inverters, Energy Meters, Pyranometers / Irradiance Sensors, Battery Storage Systems (BESS), Weather Stations, and Trackers.
- 🔍 **Interactive Register Preview**: Live filterable data table with quick actions (*Open CSV*, *Open Folder*, *Copy Data*).
- 🛡️ **Embedded Validator**: Immediate diagnostic feedback with color-coded severity badges (green = valid, orange = warning, red = blocker).

```powershell
# Commands to launch the GUI:
deffilegen-gui
deffilegen gui
launch_gui.bat
```

---

### 2. High-Performance CLI (`deffilegen`)

The command-line suite offers 10 unified subcommands:

```powershell
# 1. Complete end-to-end workflow (Extract + Generate + Validate)
deffilegen run "datasheet.pdf" --manufacturer "Huawei" --model "SUN2000-50KTL" -o huawei.csv

# 2. Batch process an entire directory of documents with auto-metadata inference
deffilegen batch datasheets_folder/ -o generated_defs/ --tag --report batch_report.md

# 3. Extract registers to an intermediate raw CSV
deffilegen extract "register_map.xlsx" --sheet "Holding Registers" -o raw_registers.csv

# 4. Generate a Webdyn definition from intermediate CSV data
deffilegen generate raw_registers.csv --manufacturer "SMA" --model "STP-5000" -o sma_def.csv

# 5. Validate an existing definition CSV for Webdyn syntax compliance
deffilegen validate sma_def.csv

# 6. Enrich or normalize official Webdyn SunPM semantic tags
deffilegen tag sma_def.csv --category Inverter

# 7. Deep semantic audit (Modbus address overlaps, tag conformity, unit coherence)
deffilegen audit Definitions/ --strict-overlap

# 8. Auto-repair definitions (normalize legacy tags, fill known tags, deduplicate)
deffilegen audit defective_def.csv --fix -o sanitized_def.csv

# 9. Generate comprehensive technical Markdown / JSON reports
deffilegen report Definitions/ -o AUDIT_REPORT.md --json metrics.json

# 10. Train semantic knowledge base from definition datasets
deffilegen train --dir Definitions/ --summary --export-json learned_knowledge.json
```

#### Helpful CLI Flags:
- `--pages 12,14-18` : Target specific register table pages in a large manual.
- `--sheet "Registers"` : Target a specific worksheet inside an Excel workbook.
- `--mapping custom.json` : Supply explicit column mapping overrides.
- `--address-offset -1` : Adjust 0-based vs 1-based numbering conventions.
- `--force` : Overwrite existing output files without prompting.
- `-v` or `--verbose` : Print granular parsing heuristics and diagnostic traces.

---

### 3. Industrial Web Studio & REST API (FastAPI Browser App)

The built-in Web Studio provides a responsive single-page application with 4 specialized engineering tabs:

1. **Convert Datasheet**: Drag-and-drop ingestion of PDF, XLSX, XLSM, CSV, or XML documents, interactive configuration (manufacturer, model, protocol, category), and real-time tabular preview with one-click CSV export.
2. **SunPM Tagger**: Upload an existing or newly generated definition to automatically enrich and align register tags with canonical Webdyn SunPM semantics (`P_AC`, `Q_AC`, `E_TOT`, etc.) with phase inference and category override.
3. **Validator & Auditor**: In-depth rule validation checking syntax, 11-column integrity, Modbus interval address collisions/overlaps, and providing a 1-click automatic remediation download.
4. **Equipment Library**: Real-time searchable catalog of 125+ pre-validated industrial definition CSV files covering 61 manufacturers, with instant live search and direct download buttons.

```powershell
uvicorn web.app:app --host 127.0.0.1 --port 8000
```
- **Web Studio Application**: [http://127.0.0.1:8000](http://127.0.0.1:8000)
- **Interactive OpenAPI Documentation**: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- **REST Endpoints**:
  - `POST /api/convert`: Multipart document conversion pipeline.
  - `POST /api/validate`: Structural and syntax validator.
  - `POST /api/tag`: SunPM semantic tag enrichment engine.
  - `POST /api/audit`: Comprehensive audit and automatic remediation engine.
  - `GET /api/library`: Equipment definition catalog metadata provider.
  - `GET /api/library/{filename}`: Path-traversal-guarded definition file download.

---

### 4. Python Programmatic API

Embed the conversion engine directly into custom Python pipelines:

```python
from generate_webdyn_def import generate_webdyn_definition

success = generate_webdyn_definition(
    input_file="manufacturer_datasheet.pdf",
    output_file="webdyn_definition.csv",
    manufacturer="SolarEdge",
    model="SE10000",
    protocol="modbusRTU",
    category="Inverter",
    address_offset=0,
    strict_validation=True,
)

if success:
    print("WebdynSunPM definition generated and validated successfully!")
```

---

## 🧠 Semantic Knowledge Miner & SunPM Tagging Engine

The **Semantic Knowledge Miner** (`DefFileGenerator/knowledge_miner.py`) is an autonomous domain intelligence layer trained on **211,000+ Modbus registers** across 250+ solar inverters, energy meters, weather stations, and battery storage systems, combined with the official Webdyn SunPM specification (`SunPM_DEfFile_Tags_v1.xlsx`).

### 1. Modern SunPM Canonical Tag Taxonomy

The engine replaces legacy inconsistent tags with official Webdyn SunPM canonical tags:

| Telemetry Domain | Canonical SunPM Tag | Legacy Equivalent(s) | Description |
|---|---|---|---|
| **Active Power** | `P_AC` | `RealPower`, `Pac-Tag`, `ActivePowSumkW` | Total active AC power (W / kW) |
| **Phase Active Power** | `P_AC_1`, `P_AC_2`, `P_AC_3` | `ActivePow1kW`, `Pac_L1` | Per-phase AC active power |
| **Reactive Power** | `Q_AC`, `Q_AC_1`..`3` | `ReactivePower`, `ReactivePowSumkW` | Total and per-phase reactive power (var / kvar) |
| **Total Energy** | `E_TOT` | `EnergyTotal`, `TotalEnergy`, `Eac-Tag` | Cumulative lifetime energy production (Wh / kWh) |
| **Daily Energy** | `E_DAY` | `EnergyDay`, `DailyEnergy`, `EacDay-Tag` | Daily energy yield (Wh / kWh) |
| **Voltages** | `U_AC_1`, `U_AC_2`, `U_AC_3` | `Vac_L1`, `GridVoltage1` | Grid phase-to-neutral / phase-to-phase voltage |
| **Currents** | `I_AC_1`, `I_AC_2`, `I_AC_3` | `Iac_L1`, `GridCurrent1` | Grid phase AC currents |
| **DC Strings** | `U_DC_1`..`N`, `I_DC_1`..`N` | `Vdc_PV1`, `Idc_PV1` | Photovoltaic MPPT string voltages and currents |
| **Frequency & PF** | `FREQ`, `PF` | `GridFrequency`, `CosPhi`, `cosPHI` | AC grid frequency (Hz) and Power Factor |
| **Cabinet Temp** | `TEMP_CAB`, `TEMP_HS` | `InternalTemperature`, `TempHeatsink` | Internal cabinet and heatsink temperatures (°C) |
| **Status & Alarms** | `STATUS`, `ALARM`, `STATE` | `InverterState`, `DeviceStatus` | Operating status and fault alarm bitfields |
| **Remote Commands** | `CMD_ON`, `CMD_OFF`, `CMD_P_AC_SET_PERCENT` | `cmdOn`, `cmdPwrPercent` | Remote Curtailment, Setpoint, and Control |

### 2. Intelligent Semantic Disambiguation

The knowledge miner incorporates multi-signal context disambiguation rules:
- **Phase Detection**: Disambiguates `L1/L2/L3`, `R/S/T`, or `Phase A/B/C` into canonical suffixes `_1`, `_2`, `_3`.
- **Absolute vs. Percentage Power**: Distinguishes active power setpoints (`CMD_P_AC_SET_ABS` in kW) from curtailment percentage commands (`CMD_P_AC_SET_PERCENT` in %).
- **Apparent vs. Reactive Power**: Accurately classifies $S$ (VA/kVA) into `S_AC` versus $Q$ (var/kvar) into `Q_AC`.
- **Automated Action Code Attribution**:
  - `4`: Telemetry sensor / read-only measurement (Holding/Input registers).
  - `9`: Status, fault alarm, or warning indicator bitfield.
  - `1`: Modbus command / control setpoint (Coils, Holding write registers).
  - `0`: Reserved / unmapped memory slot.

---

## 📊 Technical Audit & Quality Reporting Engine

The **Technical Report Generator** (`DefFileGenerator/report_generator.py`) provides automated reporting and telemetry KPI auditing across equipment definition files:

```powershell
# Generate a Markdown report for a single definition or full directory:
deffilegen report Definitions/ -o docs/RAPPORT_TECHNIQUE.md --json metrics.json
```

### Audited Quality Indicators:
1. **Structural Completeness**: 11-column strict Webdyn conformity, BOM encoding validation.
2. **Modbus Topology**: Function code split (FC01/FC02 Coils vs. FC03/FC04 Registers), memory span, address boundaries.
3. **Semantic Density**: Percentage of registers mapped to standard SunPM canonical tags.
4. **Action Code Distribution**: Read (4) vs. Status/Alarm (9) vs. Command (1) vs. Reserve (0).
5. **Memory Overlap Analysis**: Bisect interval collision check to guarantee zero overlapping register words.

---

## 🔋 Industrial Equipment Suites (Mars Renewable & FoxESS)

DefFileGenerator includes specialized industrial generators for utility-scale solar and BESS (Battery Energy Storage Systems):

### 1. Mars Renewable BESS Local Controller V7.0 (`DefFileGenerator/mars_generator.py`)
Generates standardized Webdyn definition architectures for multi-megawatt battery containers:
- **Slave 1 (Local Controller Base)**: Liquid cooling, HVAC, PCS, and fire suppression telemetry.
- **Slave 247 (EMS Energy Management System)**: System overview read telemetry (FC04) and EMS power dispatch writes (FC03).
- **Slave N (Rack & Cell Hierarchy)**: High-granularity monitoring of battery racks and up to 1,000 individual cell voltages/temperatures.
- **All-in-One Generator**: Complete unified definition linking PV, BESS, Inverters, and Substation Meters.
- Full Report: [`docs/RAPPORT_TECHNIQUE_MARS_RENEWABLE.md`](docs/RAPPORT_TECHNIQUE_MARS_RENEWABLE.md).

### 2. FoxESS P3 & H3 Pro Hybrid Storage Inverters (`scratch/regenerate_foxess_p3_h3_pro.py`)
Generates production-grade definition mappings for three-phase commercial hybrid inverters:
- Full Report: [`docs/RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md`](docs/RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md).

---

## 🏛️ Apollo-Standard Engineering Architecture & Ledgers

This repository is governed by the **Apollo-Standard Autonomous Constitution** (`AGENTS.md` and `MASTER_CODEX_MISSION.md`). Engineering changes adhere to strict evidence-based invariants:

> [!IMPORTANT]
> **Règle absolue de sécurité réseau :**  
> Toujours demander confirmation explicite à l'utilisateur avant d'écrire, modifier, écraser ou supprimer tout fichier situé sur le partage réseau commun (`\\WDN-FILES.WEBDYN.COM`).

### Durable Engineering Audit Ledgers (`docs/audit/`):
- 📋 **[`docs/audit/PLAN.md`](docs/audit/PLAN.md)**: Master engineering mission plan, phases, and milestone verification.
- 🗺️ **[`docs/audit/REVIEW_MATRIX.md`](docs/audit/REVIEW_MATRIX.md)**: Complete file inventory (97 source modules, 22,500+ LOC), domain hazards, and test coverage.
- 🔍 **[`docs/audit/FINDINGS.md`](docs/audit/FINDINGS.md)**: Traceable defect ledger (FIND-001 through FIND-003) with root causes and closure evidence.
- 🛡️ **[`docs/audit/VALIDATION.md`](docs/audit/VALIDATION.md)**: Gate results across linting, strict typing, security, build, and test coverage.
- ⚡ **[`docs/audit/PERFORMANCE.md`](docs/audit/PERFORMANCE.md)**: Benchmarks documenting an 81.3% test collection acceleration (2.50s) and memory profiling.
- 🤝 **[`docs/audit/HANDOFF.md`](docs/audit/HANDOFF.md)**: Operational handoff, reproducible commands, and production maintenance instructions.

### Cross-Platform Bootstrap Scripts:
- **Windows**: [`scripts/bootstrap.ps1`](scripts/bootstrap.ps1) — sets up uv environment, validates gates, and verifies CLI.
- **Linux/macOS**: [`scripts/bootstrap.sh`](scripts/bootstrap.sh) — sets up virtualenv, installs package, and runs verification.

---

## 📊 System Architecture & Pipeline

```text
┌────────────────────────────────────────────────────────┐
│  Source Input: PDF, Excel (.xlsx), CSV, XML            │
└───────────────────────────┬────────────────────────────┘
                            │ Safe memory buffer ingestion
                            ▼
┌────────────────────────────────────────────────────────┐
│  1. MULTI-FORMAT EXTRACTOR (extractor.py)              │
│  - Engine parsers (pdfplumber, openpyxl, defusedxml)   │
│  - Multi-row banner header merging                     │
│  - 4-tier heuristic column mapping                     │
└───────────────────────────┬────────────────────────────┘
                            │ Streaming intermediate register dicts
                            ▼
┌────────────────────────────────────────────────────────┐
│  2. SEMANTIC KNOWLEDGE MINER (knowledge_miner.py)      │
│  - 211,000+ registers training dataset                 │
│  - Canonical SunPM tag attribution (P_AC, E_TOT...)    │
│  - Action code inference (4=Telemetry, 9=Alarm...)     │
│  - Phase & scaling factor disambiguation               │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│  3. NORMALIZATION ENGINE (def_gen.py)                  │
│  - Address resolution (Hex 0x..., Decimal, Ranges)     │
│  - Data type standardization (U16, I32, F32, BITS...)  │
│  - Polynomial scaling computation (CoefA, CoefB)       │
│  - Unique variable tag generation                      │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│  4. STRICT VALIDATION LAYER (def_gen.py)               │
│  - Address boundary checks (0..65535)                  │
│  - Bisect interval lookup and overlap detection       │
│  - BITS compound validation (startbit 0..15, len 1..16)│
│  - CSV formula injection escaping                      │
└───────────────────────────┬────────────────────────────┘
                            │ Atomic write (fsync + replace)
                            ▼
┌────────────────────────────────────────────────────────┐
│  Output: WebdynSunPM Definition CSV                    │
│  (UTF-8 with BOM, semicolon-delimited, 11 columns)     │
└────────────────────────────────────────────────────────┘
```

---

## 📑 Supported Input Formats & Modbus Semantics

### 1. Accepted Input Formats
- **PDF** : Text-based tabular pages parsed via `pdfplumber`. *(Scanned image-only PDFs require external OCR).*
- **Excel** : Modern `.xlsx`, `.xlsm`, `.xltx`, `.xltm` files handled via `openpyxl`. *(Legacy 1997 `.xls` files must be saved as `.xlsx`).*
- **CSV** : Automatic delimiter detection (comma, semicolon, tab) and multi-encoding recognition (UTF-8, UTF-8-SIG, CP1252, Latin-1).
- **XML** : Hardened XML parsing with `defusedxml` to block XXE attacks.

### 2. Recognized Column Headers
The extractor recognizes headers in English, French, and German:

| Target Field | Purpose | Sample Vendor Headers Recognized |
|---|---|---|
| **Address** | Register start address | `address`, `addr`, `register`, `reg`, `registre`, `adresse`, `offset`, `index` |
| **Name** | Human description | `name`, `description`, `parameter`, `variable`, `signal`, `grandeur`, `nom` |
| **Type** | Modbus data type | `type`, `data type`, `datatype`, `format`, `taille`, `longueur` |
| **Unit** | Engineering unit | `unit`, `units`, `unite`, `unité`, `symbole` |
| **Factor** | Multiplier / Scale | `scale`, `factor`, `multiplier`, `ratio`, `gain`, `pas`, `coeff`, `multiplicateur` |
| **Offset** | Addition bias | `offset`, `bias`, `decalage`, `b` |
| **Action** | Access permission | `action`, `access`, `read/write`, `r/w`, `droits` |

### 3. Normalized Data Types

| Canonical Code | Common Vendor Names Recognized | Modbus Format |
|---|---|---|
| **U16** | `uint16`, `unsigned short`, `hex16`, `word`, `int16u` | 1 Modbus word (16-bit unsigned) |
| **I16** | `int16`, `short`, `signed short`, `int16s` | 1 Modbus word (16-bit signed) |
| **U32** | `uint32`, `unsigned long`, `dword`, `int32u`, `datetime` | 2 Modbus words (32-bit unsigned) |
| **I32** | `int32`, `long`, `signed long`, `int32s` | 2 Modbus words (32-bit signed) |
| **F32** | `float`, `float32`, `single`, `32-bit IEEE 754` | 2 Modbus words (single-precision float) |
| **U64 / I64** | `uint64`, `int64`, `qword` | 4 Modbus words (64-bit integer) |
| **F64** | `double`, `float64`, `64-bit IEEE 754` | 4 Modbus words (double-precision float) |
| **BITS** | `bit16`, `bitmap`, `bitfield`, e.g. `30001_0_1` | Bitfield slice within a 16-bit register |
| **STRING** | `string 20`, `str*30`, e.g. `30030_20` | ASCII character string |
| **RAW** | `raw`, e.g. `30040_10` (even byte length) | Uninterpreted raw bytes sequence |

---

## 🚨 Beginner SOS & Troubleshooting FAQ

Encountering an issue? Find your immediate solution here:

| Error Message or Symptom | Probable Cause | Immediate Fix |
|---|---|---|
| **`python` is not recognized...** | Python was not added to your system PATH variable during installation. | Try typing `py` instead of `python`. If that fails, reinstall Python and make sure to check **Add python.exe to PATH**. |
| **Running scripts is disabled on this system...** | Default Windows PowerShell execution security policy. | Run this command in PowerShell: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` then re-run activation. |
| **Application closes immediately when opening `.exe`** | The executable was double-clicked from inside the `.zip` without extraction. | Right-click the ZIP archive > **Extract All...**, and launch the `.exe` from the unzipped folder. |
| **Blue screen: "Windows protected your PC"** | Standard Windows Defender SmartScreen notice for new open-source software. | Click **More info**, then click **Run anyway**. |
| **No registers extracted (0 registers found)** | The PDF contains scanned pictures of tables rather than selectable text. | Run the PDF through an OCR tool (e.g. Adobe Acrobat or an online OCR tool) to make text selectable, then retry. |
| **All register addresses are shifted by +1 or -1** | The manufacturer used 0-based indexing instead of 1-based indexing (or vice versa). | Add `--address-offset 1` or `--address-offset -1` to your command to re-align all addresses in bulk. |
| **Address overlap validation error** | Multiple variables occupy the same Modbus memory words. | Check the manual and choose the intended variables. `deffilegen validate FILE --lenient` permits overlap diagnostics; normal generation remains strict. |
| **Web browser does not open automatically** | The uvicorn server started, but terminal commands do not launch browsers automatically. | Open your browser manually and navigate to: `http://127.0.0.1:8000`. |
| **`File already exists`** | The output CSV filename is already taken on your disk. | Pick a different destination name or pass the `--force` flag to permit overwriting. |

---

## 🛡️ Security & Data Integrity Model

DefFileGenerator treats all external documentation files as **untrusted input**:

- **CSV Formula Injection Mitigation (DDE Defense)**: Any cell beginning with active formula triggers (`=`, `+`, `-`, `@`, `|`, `%`, or fullwidth Unicode variants) is prefixed with an apostrophe `'` to prevent arbitrary spreadsheet code execution.
- **XML Entity & XXE Defense**: XML extraction relies exclusively on `defusedxml.ElementTree`. External DTDs, external entity expansions, and "Billion Laughs" denial-of-service payloads are strictly rejected.
- **Atomic File Writes**: Output files are staged into hidden sibling temporary files, flushed with `os.fsync`, and replaced with `os.replace`. Handled failures preserve the previous destination and remove staging files.
- **Upload Bounding**: Web uploads are streamed in 1 MiB chunks, capped at 10 MiB, and restricted to 65,536 registers. Excel ZIP expansion and worksheet dimensions are bounded before expensive processing.
- **Destination Protection**: Conversion validates a staged definition before replacing an existing destination. Source documents are protected even when the output path is a symbolic-link or hard-link alias.

See [`docs/security.md`](docs/security.md) for the complete security specification.

---

## 🧪 Quality Gates & Automated Tests

The codebase is protected by automated unit and integration tests under strict Apollo standards:

```powershell
# 1. Run the entire test suite with branch coverage (1,040 tests, 88% coverage):
uv run pytest --cov=DefFileGenerator --cov=web

# 10. Check code style and formatting:
uv run ruff check .
uv run ruff format --check .

# 10. Static type verification (Strict Mypy):
uv run mypy DefFileGenerator web

# 10. Security vulnerability scan:
uv run bandit -r DefFileGenerator web -ll

# 10. Build distribution package:
uv build --offline
```

---

## 📚 Full Documentation Index

| Guide | Description |
|---|---|
| 📖 **[`QUICKSTART.md`](QUICKSTART.md)** | Practical walkthrough with real-world manufacturer file conversions (GoodWe, SMA, ABB). |
| 📋 **[`docs/input-format.md`](docs/input-format.md)** | Technical reference for input columns, BITS compound rules, and data types. |
| 🏛️ **[`docs/architecture.md`](docs/architecture.md)** | Detailed architecture flowcharts, component invariants, and public interfaces. |
| 🛡️ **[`docs/security.md`](docs/security.md)** | In-depth security model, CSV injection defense, and automation protections. |
| 📑 **[`docs/WEBDYN_TAGS_REFERENCE.md`](docs/WEBDYN_TAGS_REFERENCE.md)** | Official SunPM tag dictionary & supervisory telemetry reference. |
| 🔋 **[`docs/RAPPORT_TECHNIQUE_MARS_RENEWABLE.md`](docs/RAPPORT_TECHNIQUE_MARS_RENEWABLE.md)** | Technical report on Mars Renewable BESS Local Controller V7.0 definition generation. |
| ⚡ **[`docs/RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md`](docs/RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md)** | Technical report on FoxESS P3/H3 Pro Inverters & Storage integration. |
| 📊 **[`docs/RAPPORT_AUDIT_DOCUMENTS_TECHNIQUES_WEBDYN.md`](docs/RAPPORT_AUDIT_DOCUMENTS_TECHNIQUES_WEBDYN.md)** | Detailed audit of technical vendor documentation conversions. |
| 📋 **[`docs/audit/PLAN.md`](docs/audit/PLAN.md)** | Apollo engineering plan, phases, and milestone tracking. |
| 🗺️ **[`docs/audit/REVIEW_MATRIX.md`](docs/audit/REVIEW_MATRIX.md)** | Codebase file review matrix and verification mapping. |
| 🔍 **[`docs/audit/FINDINGS.md`](docs/audit/FINDINGS.md)** | Durable defect and remediation ledger (FIND-001 through FIND-003). |
| 🛡️ **[`docs/audit/VALIDATION.md`](docs/audit/VALIDATION.md)** | Verification gate results, commands, and outputs. |
| ⚡ **[`docs/audit/PERFORMANCE.md`](docs/audit/PERFORMANCE.md)** | Performance benchmarks and test collection optimization. |
| 🤝 **[`docs/audit/HANDOFF.md`](docs/audit/HANDOFF.md)** | Operational handoff and maintenance documentation. |
| 📜 **[`AGENTS.md`](AGENTS.md)** | Apollo-Standard Autonomous Engineering Constitution. |
| 🎯 **[`MASTER_CODEX_MISSION.md`](MASTER_CODEX_MISSION.md)** | Full Codex Apollo mission contract and authority mandate. |
| 📝 **[`CHANGELOG.md`](CHANGELOG.md)** | Chronological history of releases, performance improvements, and fixes. |

---

## 📂 Repository Structure

```text
DefFileGenerator/
├── DefFileGenerator/         # Core Python package
│   ├── def_gen.py            # Modbus normalization, validation & CSV generation
│   ├── extractor.py          # Multi-format parser (PDF, Excel, CSV, XML)
│   ├── gui.py                # Desktop GUI application (CustomTkinter / Tkinter)
│   ├── io_utils.py           # Atomic file IO and preview utilities
│   ├── knowledge_miner.py    # Semantic SunPM tag inference (211,000+ registers dataset)
│   ├── mars_generator.py     # Mars Renewable BESS & EMS multi-slave generator
│   ├── report_generator.py   # Technical audit & quality reporting engine
│   ├── main.py               # Unified CLI entrypoint (deffilegen)
│   └── tests/                # Comprehensive test suite (1,040 tests, 88% coverage)
├── web/                      # FastAPI Web Application & UI
│   ├── app.py                # REST API endpoints (/api/convert, /api/validate)
│   └── static/               # Responsive Vanilla JS / CSS frontend
├── Definitions/              # Validated industrial equipment definitions library (125 models)
├── generated_definitions/    # Production Webdyn CSV definitions (34+ models)
├── tagged_definitions/       # Canonical SunPM-tagged production definitions
├── stress_test_data/         # Real-world vendor datasets for stress testing
├── docs/                     # Full technical documentation & audit reports
│   ├── architecture.md       # Architecture & flowcharts
│   ├── WEBDYN_TAGS_REFERENCE.md # Official SunPM tag dictionary & supervisory telemetry reference
│   ├── AUDIT_REPORT.md       # Security & code audit evidence
│   ├── RAPPORT_AUDIT_DOCUMENTS_TECHNIQUES_WEBDYN.md # Technical audit report
│   ├── RAPPORT_TECHNIQUE_MARS_RENEWABLE.md # Mars Renewable BESS technical report
│   ├── RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md # FoxESS P3/H3 Pro technical report
│   ├── input-format.md       # Input format reference
│   ├── security.md           # Security policy and mitigation models
│   ├── quickstart.md         # Quick start cheatsheet
│   └── audit/                # Durable Apollo engineering audit ledgers
│       ├── PLAN.md           # Engineering plan & milestones
│       ├── REVIEW_MATRIX.md  # Module LOC, hazards & verification matrix
│       ├── FINDINGS.md       # Defect tracking & resolution evidence
│       ├── VALIDATION.md     # Gate verification & coverage ledger
│       ├── PERFORMANCE.md    # Performance benchmarks & optimization data
│       └── HANDOFF.md        # Maintenance & operational handoff
├── scripts/                  # Cross-platform environment bootstrap scripts
│   ├── bootstrap.ps1         # Windows PowerShell setup & gate verification
│   └── bootstrap.sh          # Linux/macOS bash setup & gate verification
├── AGENTS.md                 # Apollo-Standard Autonomous Constitution
├── CLAUDE.md                 # Agent onboarding pointing to AGENTS.md
├── MASTER_CODEX_MISSION.md   # Universal Codex mission specification
├── doc_to_webdyn.py          # Standalone convenience CLI wrapper
├── generate_webdyn_def.py    # Programmatic generator API wrapper
├── build_exe.py              # PyInstaller executable builder
├── pyproject.toml            # Project packaging & tool configuration
├── DEVELOPER_GUIDE.md        # Architecture and developer manual
├── QUICKSTART.md             # Beginner walkthrough
└── README.md                 # Primary project overview
```

---

## 📄 License

This project is licensed under the open-source **MIT License**. You are free to use, modify, and distribute it in personal and commercial environments.
