# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for URL Batch Downloader
# Build command:  pyinstaller url_downloader.spec

block_cipher = None

a = Analysis(
    ['url_downloader.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # pandas / openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.workbook',
        'openpyxl.reader.excel',
        # requests
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',
        # office365-rest-python-client (loaded lazily, must be explicit)
        'office365',
        'office365.sharepoint',
        'office365.sharepoint.client_context',
        'office365.runtime',
        'office365.runtime.auth',
        'office365.runtime.auth.user_credential',
        # msal (dependency of office365)
        'msal',
        'msal.application',
        'msal.token_cache',
        # PyQt5
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.sip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'numpy',
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
