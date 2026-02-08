# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path(__file__).resolve().parent
icon_file = project_root / "Diffviewer.ico"

a = Analysis(
    ["main.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        ("reporters/templates", "reporters/templates"),
        (str(icon_file), "."),
    ],
    hiddenimports=[
        "core.qa_verify",
        "parsers.xliff_parser",
        "parsers.sdlxliff_parser",
        "parsers.memoq_parser",
        "parsers.txt_parser",
        "parsers.srt_parser",
        "parsers.xlsx_parser",
        "parsers.xls_parser",
        "parsers.pptx_parser",
        "parsers.docx_parser",
        "reporters.html_reporter",
        "reporters.excel_reporter",
        "reporters.docx_reporter",
        "reporters.summary_reporter",
        "ui.qa_column_mapping_dialog",
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
    name="Diffviewer",
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
    icon=str(icon_file),
)
