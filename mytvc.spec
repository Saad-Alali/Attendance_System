# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
    ('main-qr_attendance.py', '.'),
    ('main-Grad.py', '.'),
    ('main-scraping.py', '.'),
    ('main-finished.py', '.'),
    ('qr_attendance', 'qr_attendance'),
    ('project', 'project'),
    ('scraping', 'scraping'),
    ('security', 'security'),
    # احذف السطر التالي لأن الملف غير موجود
    # ('app_icon.ico', '.'),
]

a = Analysis(
    ['main_ui.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'flask', 'selenium', 'pandas', 'numpy', 'qrcode', 
        'pyzbar', 'openpyxl', 'PIL', 'cv2', 'difflib',
        'flask.templating', 'flask.scaffold', 'email.mime.text'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='MyTVTC',
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
    # احذف هذا السطر أو اجعله تعليقًا
    # icon='app_icon.ico',
)