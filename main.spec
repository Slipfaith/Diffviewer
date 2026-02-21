# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


spec_root = globals().get("SPECPATH")
if not spec_root:
    spec_file = globals().get("SPEC")
    if spec_file:
        spec_root = str(Path(spec_file).resolve().parent)
if not spec_root:
    raise RuntimeError("SPECPATH is not defined. Run PyInstaller with this spec file.")

project_root = Path(spec_root).resolve()
templates_dir = project_root / "reporters" / "templates"
icon_file = project_root / "Diffviewer.ico"

if not templates_dir.is_dir():
    raise FileNotFoundError(f"Templates directory not found: {templates_dir}")
if not icon_file.is_file():
    raise FileNotFoundError(f"Icon file not found: {icon_file}")

datas = [
    (str(templates_dir), "reporters/templates"),
    (str(icon_file), "."),
]

a = Analysis(
    ["main.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[
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
