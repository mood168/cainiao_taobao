# -*- mode: python ; coding: utf-8 -*-
import os

# 檢查圖標文件是否存在
icon_path = 'NONE'
if os.path.exists('icon.ico'):
    icon_path = 'icon.ico'

a = Analysis(
    ['get_cainiao_page.py'],
    pathex=[],
    binaries=[],
    datas=[('chromedriver.exe', '.')],
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
    name='菜鳥工單自動處理_v12',
    debug=True,  # 啟用除錯模式
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 顯示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
) 