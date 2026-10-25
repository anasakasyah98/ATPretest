# -*- mode: python ; coding: utf-8 -*-
import os
import sys

# Executable icon: the eagle emblem cut from EagleWireless_Logo.jpg
sys.path.insert(0, SPECPATH)
from gui import load_logo_emblem, ICON_SIZES

icon_file = None
emblem = load_logo_emblem()
if emblem is not None:
    os.makedirs(workpath, exist_ok=True)
    icon_file = os.path.join(workpath, 'EaglePretest.ico')
    emblem.save(icon_file, sizes=[(size, size) for size in ICON_SIZES])

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[('EagleWireless_Logo.jpg', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy'],      # Optional for Pillow, not used by the tool (adds ~12 MB)
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='EaglePretest',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file,
)
