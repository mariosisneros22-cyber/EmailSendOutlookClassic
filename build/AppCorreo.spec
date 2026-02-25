# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hidden = [
    'win32com.client',
    'win32com.gen_py',
    'pythoncom',
]

a = Analysis(
    [r'..\src\app.py'],
    pathex=[r'..\src'],
    binaries=[],
    datas=[(r'..\src\logo', 'logo')],
    hiddenimports=[
        'win32com.client',
        'win32com.gen_py',
        'pythoncom',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AppCorreo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI (sin consola)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='AppCorreo',
)