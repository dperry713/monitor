# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['tool.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PyQt5.QtWidgets',
        'PyQt5.QtCore', 
        'PyQt5.QtGui',
        'pyqtgraph',
        'numpy',
        'pandas',
        'obd',
        'serial',
        'serial.tools.list_ports',
        'warnings',
        'time',
        'random',
        'sys',
        'winreg',
        'subprocess',
        're',
        'platform',
        'csv'
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
    name='OBD_Monitor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Disable console for clean windowed app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
