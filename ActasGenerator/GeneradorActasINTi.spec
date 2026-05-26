# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\iniciar.py'],
    pathex=['C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\src'],
    binaries=[],
    datas=[('C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\src', 'src'), ('C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\assets', 'assets'), ('C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\config.json', '.')],
    hiddenimports=['config', 'database', 'excel_processor', 'document_generator', 'main_app', 'config_editor', 'ttkbootstrap', 'openpyxl', 'docx', 'pandas', 'PIL'],
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
    name='GeneradorActasINTi',
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
    icon=['C:\\Users\\usuario\\Documents\\Desarrollo\\ATASIS\\ActasGenerator\\assets\\logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GeneradorActasINTi',
)
