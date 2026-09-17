# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller specification file for WebdynSunPM DefFileGenerator.

Builds standalone Windows 64-bit executables:
1. DefFileGenerator-GUI.exe: Windowed desktop GUI application (customtkinter/Tkinter)
2. deffilegen.exe: High-performance console CLI binary

Usage via PyInstaller:
    pyinstaller DefFileGenerator.spec

Or use the automated orchestration script:
    python build_exe.py --target all
"""

import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Ensure the repository root is properly resolved
REPO_ROOT = os.path.abspath(SPECPATH)
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

# Determine the build target from environment variable (default: "all")
build_target = os.environ.get("BUILD_TARGET", "all").lower().strip()

# Collect dynamic assets, fonts, themes, and hidden imports for critical dependencies
packages_to_collect = [
    "customtkinter",
    "openpyxl",
    "pdfplumber",
    "pypdfium2",
    "defusedxml",
]

collected_datas = []
collected_binaries = []
collected_hiddenimports = [
    "DefFileGenerator",
    "DefFileGenerator.def_gen",
    "DefFileGenerator.extractor",
    "DefFileGenerator.gui",
    "DefFileGenerator.main",
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "darkdetect",
]

for package_name in packages_to_collect:
    try:
        pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(package_name)
        collected_datas.extend(pkg_datas)
        collected_binaries.extend(pkg_binaries)
        collected_hiddenimports.extend(pkg_hiddenimports)
    except Exception as exc:  # pragma: no cover
        print(f"[SPEC WARNING] collect_all failed for {package_name}: {exc}")

# Deduplicate collected assets
datas_dedup = list({(src, dst): (src, dst) for src, dst in collected_datas}.values())
binaries_dedup = list({(src, dst): (src, dst) for src, dst in collected_binaries}.values())
hiddenimports_dedup = sorted(set(collected_hiddenimports))

# Common exclusions to keep binary size lean and optimized
common_excludes = [
    "matplotlib",
    "scipy",
    "notebook",
    "IPython",
    "pandas",
    "pytest",
    "unittest",
    "sqlite3",
]

# ==============================================================================
# 1. DefFileGenerator-GUI Target (Windows Desktop Graphical Application)
# ==============================================================================
if build_target in ("all", "gui"):
    a_gui = Analysis(
        [os.path.join(REPO_ROOT, "DefFileGenerator", "gui.py")],
        pathex=[REPO_ROOT],
        binaries=binaries_dedup,
        datas=datas_dedup,
        hiddenimports=hiddenimports_dedup,
        hookspath=[],
        hooksconfig={},
        runtime_hooks=[],
        excludes=common_excludes,
        win_no_prefer_redirects=False,
        win_private_assemblies=False,
        cipher=block_cipher,
        noarchive=False,
    )
    pyz_gui = PYZ(a_gui.pure, a_gui.zipped_data, cipher=block_cipher)
    exe_gui = EXE(
        pyz_gui,
        a_gui.scripts,
        a_gui.binaries,
        a_gui.zipfiles,
        a_gui.datas,
        [],
        name="DefFileGenerator-GUI",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,  # Windowed GUI application (no console popup)
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )

# ==============================================================================
# 2. deffilegen Target (Command-Line Interface Binary)
# ==============================================================================
if build_target in ("all", "cli"):
    a_cli = Analysis(
        [os.path.join(REPO_ROOT, "DefFileGenerator", "main.py")],
        pathex=[REPO_ROOT],
        binaries=binaries_dedup,
        datas=datas_dedup,
        hiddenimports=hiddenimports_dedup,
        hookspath=[],
        hooksconfig={},
        runtime_hooks=[],
        excludes=common_excludes,
        win_no_prefer_redirects=False,
        win_private_assemblies=False,
        cipher=block_cipher,
        noarchive=False,
    )
    pyz_cli = PYZ(a_cli.pure, a_cli.zipped_data, cipher=block_cipher)
    exe_cli = EXE(
        pyz_cli,
        a_cli.scripts,
        a_cli.binaries,
        a_cli.zipfiles,
        a_cli.datas,
        [],
        name="deffilegen",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=True,  # Console application
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
