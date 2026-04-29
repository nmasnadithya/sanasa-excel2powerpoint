# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for sansa-excel2pptx.

Build with:
    pyinstaller sansa-excel2pptx.spec

Produces a single-file --windowed exe at dist/sansa-excel2pptx.exe.
The template (labalaba ginuma.pptx) is NOT bundled — it ships alongside
the .exe so users can swap in a customised version.
"""

block_cipher = None

a = Analysis(
    ['src/__main__.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'matplotlib.backends.backend_agg',
        'matplotlib.backends.backend_svg',
        'PIL._tkinter_finder',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='sansa-excel2pptx',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # --windowed
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
