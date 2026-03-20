# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['KarvOps.py'],
    pathex=[],
    binaries=[],
    datas=[('dashboard', 'dashboard'), ('dynatrace_tracker', 'dynatrace_tracker'), ('templates', 'templates'), ('static', 'static'), ('staticfiles', 'staticfiles'), ('db.sqlite3', '.')],
    hiddenimports=['django', 'apscheduler', 'django_apscheduler', 'requests', 'numpy', 'matplotlib', 'matplotlib.pyplot', 'matplotlib.dates', 'matplotlib.ticker', 'reportlab', 'reportlab.lib.colors', 'reportlab.lib.pagesizes', 'reportlab.platypus', 'pillow', 'PIL', 'PIL._imaging', 'openpyxl', 'openpyxl.chart'],
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
    name='KarvOps',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
