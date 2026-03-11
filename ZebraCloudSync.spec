# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import customtkinter, os

# Incluir todos los archivos de temas/assets de customtkinter
ctk_datas = collect_data_files("customtkinter")

a = Analysis(
    ['agente_zebra_cloud_sync.py'],
    pathex=[],
    binaries=[],
    datas=ctk_datas,
    hiddenimports=[
        # pystray backend Windows
        'pystray._win32',
        # PIL / Pillow
        'PIL._tkinter_finder',
        # pyserial
        'serial.tools.list_ports',
        'serial.tools.list_ports_windows',
        # oauth2client
        'oauth2client.service_account',
        'oauth2client.crypt',
        # gspread + dependencias
        'gspread',
        'gspread.exceptions',
        'google.auth',
        'google.auth.transport.requests',
        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.alignment',
        'openpyxl.styles.font',
        'openpyxl.styles.fills',
        'openpyxl.cell._writer',
        # mysql connector 9.x – incluir ambos plugins de autenticacion
        'mysql.connector',
        'mysql.connector.plugins',
        'mysql.connector.plugins.mysql_native_password',
        'mysql.connector.plugins.caching_sha2_password',
        'mysql.connector.connection_cext',
        # customtkinter
        'customtkinter',
        # pynput – captura teclado HID
        'pynput',
        'pynput.keyboard',
        'pynput.mouse',
        'pynput._util',
        'pynput._util.win32',
        'pynput.keyboard._win32',
        'pynput.mouse._win32',
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
    name='ZebraCloudSync',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
