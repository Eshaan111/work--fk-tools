# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['selenium.webdriver.firefox.webdriver', 'selenium.webdriver.firefox.service', 'selenium.webdriver.firefox.options']
hiddenimports += collect_submodules('selenium.webdriver')
hiddenimports += collect_submodules('urllib3')
hiddenimports += collect_submodules('websocket')
hiddenimports += collect_submodules('trio')
hiddenimports += collect_submodules('trio_websocket')


a = Analysis(
    ['C:\\Users\\ESHAAN\\HAKUR\\work--fk-tools\\Full-LC-AUTO\\main_tab_based.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
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
    name='Full-LC-AUTO',
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
    name='Full-LC-AUTO',
)
