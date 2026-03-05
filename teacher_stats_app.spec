# -*- mode: python ; coding: utf-8 -*-

import shutil as _shutil
import platform as _platform

_pandoc_binaries = []
_pandoc_path = _shutil.which('pandoc')
if _pandoc_path:
    _pandoc_binaries.append((_pandoc_path, '.'))

a = Analysis(
    ['teacher_stats_gui.py'],
    pathex=[],
    binaries=_pandoc_binaries,
    datas=[
        ('teacher_stats_config.toml', '.'),
    ],
    hiddenimports=[
        'teacher_stats',
        'PySimpleGUI',
        'pandas',
        'matplotlib',
        'openpyxl',
        'xlrd',
        'pypinyin',
        'tomllib',
    ],
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
    name='教师科研统计分析工具',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='教师科研统计分析工具',
)

import sys as _sys
if _sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='教师科研统计分析工具.app',
        icon=None,
        bundle_identifier='com.bnugrb.teacher_stats',
    )