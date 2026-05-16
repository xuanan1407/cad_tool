# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for CAD Column Inspector Pro
# Build command: python -m PyInstaller build_installer.spec --clean

# Version Info
VERSION = '1.0.0.0'
COMPANY_NAME = 'Tran Xuan An'
FILE_DESCRIPTION = 'CAD Column Inspector Pro - Automation Tool'
INTERNAL_NAME = 'CAD_Column_Inspector_Pro'
COPYRIGHT = 'Copyright (c) 2026 Tran Xuan An'
ORIGINAL_FILENAME = 'CAD_Column_Inspector_Pro.exe'
PRODUCT_NAME = 'CAD Column Inspector Pro'

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Add README if you want to include it
        # ('README.md', '.'),
    ],
    hiddenimports=[
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32con',
        'pyautocad',
        'ezdxf',
        'pandas',
        'openpyxl',
        'tkinter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',  # Exclude unused heavy libraries
        'scipy',
        'numpy.distutils',
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
    [],
    exclude_binaries=True,
    name='CAD_Column_Inspector_Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True if you want to see console for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon='icon.ico' if you have an icon file
    version='version_info.txt',  # Windows version info
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CAD_Column_Inspector_Pro',
)
