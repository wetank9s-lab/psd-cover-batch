# -*- mode: python ; coding: utf-8 -*-
# Stage 6.5：ttkbootstrap 1.9.0 主题/本地化资源随包分发 + Per-Monitor V2 DPI 感知
import os
import ttkbootstrap
from PyInstaller.utils.win32 import winmanifest

TTB_DIR = os.path.dirname(ttkbootstrap.__file__)
THEMES_DIR = os.path.join(TTB_DIR, "themes")

# Per-Monitor V2 DPI 感知：tkinter/Tcl 8.6 需显式声明才在 125%/150% 下清晰缩放
DPI_MANIFEST_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">'
    '<application xmlns="urn:schemas-microsoft-com:asm.v3">'
    '<windowsSettings>'
    '<dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>'
    '<dpiAwareness xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">PerMonitorV2</dpiAwareness>'
    '</windowsSettings>'
    '</application>'
    '</assembly>'
)
DPI_MANIFEST = winmanifest.create_application_manifest(
    manifest_xml=DPI_MANIFEST_XML, uac_admin=False, uac_uiaccess=False
).decode("utf-8")

a = Analysis(
    ['qifang_cover_maker.py'],
    pathex=[],
    binaries=[],
    datas=[
        (THEMES_DIR, 'ttkbootstrap/themes'),
        (os.path.join(TTB_DIR, 'localization'), 'ttkbootstrap/localization'),
    ],
    hiddenimports=[
        'win32com', 'win32com.client', 'pythoncom', 'openpyxl',
        # ttkbootstrap 1.9.0（经典版）主题/工具子模块
        'ttkbootstrap', 'ttkbootstrap.themes',
        'ttkbootstrap.widgets', 'ttkbootstrap.style',
        'ttkbootstrap.window', 'ttkbootstrap.scrolled',
        'ttkbootstrap.dialogs', 'ttkbootstrap.toast',
        'ttkbootstrap.tooltip', 'ttkbootstrap.tableview',
        'ttkbootstrap.validation', 'ttkbootstrap.utility',
        'ttkbootstrap.localization',
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
    name='七方视频封面批量制作',
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
    manifest=DPI_MANIFEST,
)
