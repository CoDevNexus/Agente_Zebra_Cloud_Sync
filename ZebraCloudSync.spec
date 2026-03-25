# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['agente_zebra_cloud_sync.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'mysql.connector.plugins.mysql_native_password',
        'mysql.connector.plugins.caching_sha2_password',
        'mysql.connector.plugins.sha256_password',
        'mysql.connector.plugins.mysql_clear_password',
        'mysql.connector.locales.eng',
        'mysql.connector.locales.eng.client_error',
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
