# -*- mode: python ; coding: utf-8 -*-
import re
from pathlib import Path

from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo,
    StringFileInfo,
    StringStruct,
    StringTable,
    VarFileInfo,
    VarStruct,
    VSVersionInfo,
)

# main.py is the single source of truth for the version; read it without
# importing, so building does not pull PySide6 into the spec evaluation.
APP_VERSION = re.search(
    r'^APP_VERSION = "([^"]+)"',
    Path('main.py').read_text(encoding='utf-8'),
    re.MULTILINE,
).group(1)

version_fields = tuple(int(part) for part in (APP_VERSION.split('.') + ['0', '0', '0'])[:4])

version_resource = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=version_fields,
        prodvers=version_fields,
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo([
            StringTable('040904B0', [
                StringStruct('FileDescription', 'Cleanup tool for PDF-converted DOCX and ODT documents'),
                StringStruct('FileVersion', APP_VERSION),
                StringStruct('InternalName', 'ScanSweep'),
                StringStruct('LegalCopyright', 'MIT License'),
                StringStruct('OriginalFilename', f'ScanSweep-{APP_VERSION}.exe'),
                StringStruct('ProductName', 'ScanSweep'),
                StringStruct('ProductVersion', APP_VERSION),
            ]),
        ]),
        VarFileInfo([VarStruct('Translation', [0x409, 1200])]),
    ],
)

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('app_icon.svg', '.'), ('info_icon.svg', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name=f'ScanSweep-{APP_VERSION}',
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
    icon=['app_icon.ico'],
    version=version_resource,
)
