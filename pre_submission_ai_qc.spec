# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

block_cipher = None
root = Path.cwd()
icon_path = root / "src" / "assets" / "app_icon.ico"
datas = [
    ("config", "config"),
    ("src/assets", "src/assets"),
]
hiddenimports = [
    "keyring.backends.Windows",
    "pandas",
    "openpyxl",
    "matplotlib.backends.backend_agg",
    "reportlab",
    "yaml",
]

a = Analysis(
    ["main.py"],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PreSubmissionAIQC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path) if icon_path.exists() else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PreSubmissionAIQC",
)
