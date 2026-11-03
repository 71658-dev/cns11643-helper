# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['全字庫小幫手.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32com.client', 'pythoncom'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'setuptools', 'pkg_resources', 'cryptography', 'OpenSSL', 'unittest', 'pydoc', '_pyrepl', 'multiprocessing', 'asyncio', 'xmlrpc', 'lib2to3', 'distutils', 'pytest', 'IPython', 'matplotlib', 'packaging', 'PIL.ImageQt', 'PyQt5', 'PySide6', 'PIL._avif', 'PIL.AvifImagePlugin', 'PIL._webp', 'PIL.WebPImagePlugin'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='全字庫小幫手',
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
    name='全字庫小幫手',
)

# 授權聲明：本專案 LICENSE（GPL-3.0）與第三方元件授權全文，放在 exe 旁邊一起散布
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.join(SPECPATH, 'tools'))
from collect_licenses import write_notices
write_notices(_os.path.join(DISTPATH, '全字庫小幫手'))
