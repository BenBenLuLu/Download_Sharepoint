# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for URL Batch Downloader
# Build command:  pyinstaller url_downloader.spec --clean

block_cipher = None

# ── Collect packages that have binary extensions or dynamic imports ────────
from PyInstaller.utils.hooks import collect_all, collect_submodules

numpy_d,    numpy_b,    numpy_h    = collect_all('numpy')
pandas_d,   pandas_b,   pandas_h   = collect_all('pandas')
openpyxl_d, openpyxl_b, openpyxl_h = collect_all('openpyxl')

a = Analysis(
    ['url_downloader.py'],
    pathex=[],
    binaries=[] + numpy_b + pandas_b + openpyxl_b,
    datas=[]   + numpy_d  + pandas_d  + openpyxl_d,
    hiddenimports=(
        numpy_h + pandas_h + openpyxl_h
        + collect_submodules('numpy')
        + collect_submodules('pandas')
        + [
            # requests
            'requests',
            'urllib3',
            'certifi',
            'charset_normalizer',
            'idna',
            # office365 (lazy import)
            'office365',
            'office365.sharepoint',
            'office365.sharepoint.client_context',
            'office365.runtime',
            'office365.runtime.auth',
            'office365.runtime.auth.user_credential',
            # msal (SharePoint device-code auth)
            'msal',
            'msal.application',
            'msal.token_cache',
            # PyQt5
            'PyQt5',
            'PyQt5.QtWidgets',
            'PyQt5.QtCore',
            'PyQt5.QtGui',
            'PyQt5.sip',
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'PIL',
        'cv2',
    ],
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
    name='URL_Batch_Downloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no black console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
