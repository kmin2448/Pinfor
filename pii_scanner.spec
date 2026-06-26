# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 빌드 스펙 — 개인정보 검출 스캐너 (GUI)

빌드:
    pip install -r requirements.txt pyinstaller
    pyinstaller pii_scanner.spec

결과물:
    dist/개인정보검출스캐너.exe   (단일 실행 파일, 콘솔창 없음)
"""

from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []

# customtkinter 는 테마/에셋 데이터 파일을 함께 묶어야 정상 동작
for pkg in ("customtkinter",):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# 선택적 파서들 — 설치돼 있으면 함께 포함 (없으면 무시됨)
hiddenimports += [
    "docx", "openpyxl", "pdfplumber", "olefile", "send2trash",
]

block_cipher = None

a = Analysis(
    ["pii_scanner_gui.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="개인정보검출스캐너",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI 앱이므로 콘솔창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
