# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for P-ID IO Count Automation Tool
#
# Build with:
#   pip install pyinstaller
#   pyinstaller P-ID-IO-Count-Automation-Tool.spec
#
# The resulting executable is placed in dist/P-ID-IO-Count-Automation-Tool.exe
# (Windows) or dist/P-ID-IO-Count-Automation-Tool (Linux/macOS).

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect PySide6 Qt platform plugins and other data files that PyInstaller
# sometimes misses with the automatic analysis.
datas = []
datas += [("logo.png", ".")]          # embed logo for home screen & window icon
datas += collect_data_files("PySide6", includes=["*.dll", "*.so", "plugins/**/*"])
datas += collect_data_files("pypdf")
datas += collect_data_files("openpyxl")

hiddenimports = []
hiddenimports += collect_submodules("PySide6")
hiddenimports += collect_submodules("pypdf")
hiddenimports += collect_submodules("openpyxl")

a = Analysis(
    ["test_code.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="P-ID-IO-Count-Automation-Tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # windowed=True hides the console window on Windows so only the GUI appears.
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="logo.png",
)
