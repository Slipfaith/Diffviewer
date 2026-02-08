# Change Tracker

Change Tracker is a Windows desktop tool for deterministic comparison of localization and editor files.  
It detects `ADDED`, `DELETED`, `MODIFIED`, `UNCHANGED` changes and generates visual reports.

## Features

- File vs File comparison
- Batch folder comparison
- Multi-version chain comparison
- HTML report (interactive, self-contained)
- Excel report (rich inline diff via XlsxWriter)
- DOCX Track Changes (Microsoft Word COM, when available)

## Supported Formats

- XLIFF family: `.xliff`, `.xlf`, `.sdlxliff`, `.mqxliff`
- Text: `.txt`, `.srt`
- Office: `.xlsx`, `.xls`, `.pptx`, `.docx` (`.doc` parse fallback not fully implemented)

## Run From Source

```bash
pip install -r requirements.txt
python main.py
```

- `python main.py` without args opens GUI.
- `python main.py ...` with args runs CLI.

## CLI Examples

```bash
python cli.py compare tests/fixtures/sample_a.xliff tests/fixtures/sample_b.xliff -o ./output/
python cli.py batch folder_a/ folder_b/ -o ./output/
python cli.py versions v1.xliff v2.xliff v3.xliff -o ./output/
python cli.py formats
```

## Build Standalone EXE

```bash
pip install pyinstaller
pyinstaller change_tracker.spec
```

Output:

- `dist/ChangeTracker.exe`

You can also use:

```bash
python build.py
```

## Requirements

- Python 3.11+
- Windows
- Microsoft Word installed (for DOCX Track Changes mode)

