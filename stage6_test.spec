# -*- mode: python ; coding: utf-8 -*-
# Stage 6 测试 EXE（不正式发布）：在正式 GUI spec 基础上补充 core 包 hiddenimports

a = Analysis(
    ['qifang_cover_maker.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32com', 'win32com.client', 'pythoncom', 'openpyxl',
        'core', 'core.task_events', 'core.app_state', 'core.worker_base',
        'core.photoshop', 'core.layer_index', 'core.excel_data',
        'core.renderer', 'core.logo_mapping', 'core.output_paths', 'core.util',
        'gui_workers',
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
    a.binaries,
    a.datas,
    [],
    name='七方视频封面批量制作_stage6_test',
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
)
