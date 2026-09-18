# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import os
import platform
import struct
import sys

if sys.platform not in ('darwin', 'win32'):
    raise SystemExit('Build on Apple Silicon macOS or AMD64 Windows.')
if sys.platform == 'win32' and (struct.calcsize('P') != 8 or platform.machine().lower() not in ('amd64', 'x86_64')):
    raise SystemExit('Use AMD64 (64-bit) Python to build the Windows application.')

try:
    import tkinter
except ImportError as exc:
    raise SystemExit(
        'Tcl/Tk is required to build the GUI. On Homebrew Python 3.14, run '
        'brew install python-tk@3.14. On Windows, install Python with Tcl/Tk support.'
    ) from exc

root = Path(SPECPATH)
is_macos = sys.platform == 'darwin'
app_name = 'SpDocumentsConverter'
if is_macos:
    import runpy
    macos_compat = runpy.run_path(str(root / 'packaging/check_macos_bundle.py'))

a = Analysis(
    [str(root / 'main.py')],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / 'src/read/toggle' / name), 'src/read/toggle')
           for name in ('soo_lu.xlsx', 'happy_lu.xlsx')],  # ASCII 파일명: 한글 이름은 압축 해제 시 NFC/NFD 정규화로 서명 봉인이 깨질 수 있음
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
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=not is_macos,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='arm64' if is_macos else None,
    codesign_identity=(os.environ.get('MACOS_CODESIGN_IDENTITY') or None) if is_macos else None,
    entitlements_file=str(root / 'packaging/macos-entitlements.plist') if is_macos else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=app_name,
)

if is_macos:
    app = BUNDLE(
        coll,
        name=app_name + '.app',
        bundle_identifier='com.mirstream.spdocumentsconverter',
        info_plist={
            'LSMinimumSystemVersion': macos_compat['MIN_MACOS_VERSION'],
            'NSHighResolutionCapable': True,
            'NSAppleEventsUsageDescription': '열려 있는 Excel 문서의 시트와 선택 영역을 읽어 변환합니다.',
        },
    )
    macos_compat['check_bundle'](app.name)
