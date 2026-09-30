# WebdynSunPM DefFileGenerator & Documentation Parser

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.10+" />
  <img src="https://img.shields.io/badge/Platform-Windows%2011%20%7C%20Linux%20%7C%20macOS-0078D4?style=for-the-badge&logo=windows&logoColor=white" alt="Platforms" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="MIT License" />
  <img src="https://github.com/Cylae/DefFileGenerator/actions/workflows/ci.yml/badge.svg" alt="Continuous integration" />
  <img src="https://img.shields.io/badge/Code%20Style-Ruff-black?style=for-the-badge" alt="Ruff" />
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
- [📊 System Architecture & Pipeline](#-system-architecture--pipeline)
- [📑 Supported Input Formats & Modbus Semantics](#-supported-input-formats--modbus-semantics)
- [🚨 Beginner SOS & Troubleshooting FAQ](#-beginner-sos--troubleshooting-faq)
- [🛡️ Security & Data Integrity Model](#️-security--data-integrity-model)
- [🧪 Quality Gates & Automated Tests](#-quality-gates--automated-tests)
- [📚 Full Documentation Index](#-full-documentation-index)

---

## 🎯 What is DefFileGenerator?

To connect an industrial photovoltaic inverter (Huawei, SMA, SolarEdge, Sungrow, Fronius, ABB, etc.), energy meter, or weather station to a **WebdynSunPM** data logger, you must provide a strict semicolon-delimited CSV definition file containing every Modbus register address, data type, physical unit, and scaling factor.

Manually creating this file usually requires hours of tedious, error-prone manual work: copy-pasting tables from PDF user manuals or Excel spreadsheets, converting hexadecimal addresses to decimal numbers, calculating scale factors, and checking for overlapping memory registers.

**DefFileGenerator automates the entire process in seconds:**
1. **Reads your source document** (PDF datasheet, Excel workbook, CSV table, or XML register map).
2. **Automatically locates register tables** and understands vendor-specific column names in English, French, and German.
3. **Normalizes data types and addresses** (hex/decimal conversions, standard types like `U16`, `I32`, `F32`, individual bitfield slices).
4. **Performs strict validation** (unique variable tags, valid Modbus address spaces, $O(\log N)$ collision and overlap checks).
5. **Generates an official WebdynSunPM CSV file**, ready to be imported directly into your data logger gateway!

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
│ • You want to install and run the tool properly via Python:            │
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
   - If you see `Python 3.10.x`, `Python 3.11.x`, or `Python 3.12.x`: **Great! Skip directly to Step 2.**
   - If Windows opens the Microsoft Store or prints *"'python' is not recognized"*: **Proceed to Step 1 below.**

#### Step 1: Install Python (Watch out for the #1 beginner trap!)
1. Visit the official Python download page: [python.org/downloads](https://www.python.org/downloads/).
2. Click the yellow button **Download Python 3.12.x** (or any version $\ge$ 3.10).
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

# 2. Activate the virtual environment
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

*Installation takes about 30 seconds to fetch required libraries (pdfplumber, openpyxl, customtkinter, defusedxml).*

#### Step 6: Verify the installation!

Run:
```powershell
deffilegen --version
```
The terminal will respond:
```text
deffilegen 0.2.1
```

You are all set! 🎉

---

### Method 4 — For macOS and Linux Users

On macOS or Linux distributions (Ubuntu, Debian, Fedora, Arch):

```bash
# 1. Clone the repository
git clone https://github.com/Cylae/DefFileGenerator.git
cd DefFileGenerator

# 2. Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 3. Install dependencies (CLI and Web interface)
pip install --upgrade pip
pip install -e ".[web]"

# 4. Verify installation
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
   `sample_inverter_registers.xlsx`. You can also create a CSV template with `deffilegen template template.csv`.
4. Notice that **Manufacturer** and **Model** fields auto-fill if detected from the filename.
5. Set your destination CSV file path (or keep the suggested default).
6. Click the green **Generate Definition File** button.
7. When conversion completes, inspect the logs and the preview of the generated registers. Processing time depends on the source document.
8. Click the **Validator** tab to run instant verification on your generated file.

---

### Case B: Using the Command Line Interface (CLI)

Open your terminal in the project directory and run your first conversion:

```powershell
# Direct conversion of the bundled sample Excel register map
deffilegen run sample_inverter_registers.xlsx --manufacturer "MyBrand" --model "Inverter-5000" -o my_inverter.csv

# Immediate audit of the generated CSV file
deffilegen validate my_inverter.csv
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
| **Python Core API** | Integrators & Developers | Clean functional Python API for data ingestion pipelines | `import generate_webdyn_def` |

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

The command-line suite offers 4 primary subcommands:

```powershell
# 1. Complete end-to-end workflow (Extract + Generate + Validate)
deffilegen run "datasheet.pdf" --manufacturer "Huawei" --model "SUN2000-50KTL" -o huawei.csv

# 2. Extract registers to an intermediate raw CSV
deffilegen extract "register_map.xlsx" --sheet "Holding Registers" -o raw_registers.csv

# 3. Generate a Webdyn definition from intermediate CSV data
deffilegen generate raw_registers.csv --manufacturer "SMA" --model "STP-5000" -o sma_def.csv

# 4. Audit an existing definition CSV for WebdynSunPM compliance
deffilegen validate sma_def.csv
```

#### Helpful CLI Flags:
- `--pages 12,14-18` : Target specific register table pages in a large manual.
- `--sheet "Registers"` : Target a specific worksheet inside an Excel workbook.
- `--mapping custom.json` : Supply explicit column mapping overrides.
- `--address-offset -1` : Adjust 0-based vs 1-based numbering conventions.
- `--force` : Overwrite existing output files without prompting.
- `-v` or `--verbose` : Print granular parsing heuristics and diagnostic traces.

---

### 3. Local Web Interface (FastAPI Browser App)

```powershell
uvicorn web.app:app --host 127.0.0.1 --port 8000
```
- Web Application: [http://127.0.0.1:8000](http://127.0.0.1:8000)
- Interactive OpenAPI Docs: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- Endpoints: `POST /api/convert` (file conversion) and `POST /api/validate` (compliance check).

---

### 4. Python Programmatic API

Embed the conversion engine directly into custom Python tools:

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
│  2. NORMALIZATION ENGINE (def_gen.py)                  │
│  - Address resolution (Hex 0x..., Decimal, Ranges)     │
│  - Data type standardization (U16, I32, F32, BITS...)  │
│  - Polynomial scaling computation (CoefA, CoefB)       │
│  - Unique variable tag generation                      │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│  3. STRICT VALIDATION LAYER (def_gen.py)               │
│  - Address boundary checks (0..65535)                  │
│  - O(log N) bisect register overlap detection          │
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
- **Atomic File Writes**: Output files are staged into hidden sibling temporary files, synced to physical storage with `os.fsync`, and atomically renamed with `os.replace`. Partial or corrupted files are never left on disk.
- **Upload Bounding**: Web uploads are streamed in 1 MiB chunks, capped at 10 MiB, and restricted to 65,536 registers. Excel ZIP expansion and worksheet dimensions are bounded before expensive processing.
- **Destination Protection**: Conversion validates a staged definition before replacing an existing destination. Source documents are protected even when the output path is a symbolic-link or hard-link alias.

See [`docs/security.md`](docs/security.md) for the complete security specification.

---

## 🧪 Quality Gates & Automated Tests

The codebase is protected by automated unit and integration tests:

```powershell
# Run the entire test suite
pytest

# Check code formatting and linting
ruff check .
ruff format --check .

# Static type verification
mypy DefFileGenerator web
```

---

## 📚 Full Documentation Index

| Guide | Description |
|---|---|
| 📖 **[`QUICKSTART.md`](QUICKSTART.md)** | Practical walkthrough with real-world manufacturer file conversions (GoodWe, SMA, ABB). |
| 📋 **[`docs/input-format.md`](docs/input-format.md)** | Technical reference for input columns, BITS compound rules, and data types. |
| 🏛️ **[`docs/architecture.md`](docs/architecture.md)** | Detailed architecture flowcharts, component invariants, and public interfaces. |
| 🛡️ **[`docs/security.md`](docs/security.md)** | In-depth security model, CSV injection defense, and automation protections. |
| 💻 **[`DEVELOPER_GUIDE.md`](DEVELOPER_GUIDE.md)** | Onboarding manual and extension recipes for engineers working on the codebase. |
| 🛠️ **[`docs/development.md`](docs/development.md)** | Developer workflow, Git guidelines, and static analysis checklists. |
| 📊 **[`AUDIT_REPORT.md`](AUDIT_REPORT.md)** | Engineering audit report, test metrics, and hardening history. |
| 📝 **[`CHANGELOG.md`](CHANGELOG.md)** | Chronological history of releases, performance improvements, and fixes. |

---

## 📄 License

This project is licensed under the open-source **MIT License**. You are free to use, modify, and distribute it in personal and commercial environments.
