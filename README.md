# Diff View

`Diff View` is a Windows desktop tool for deterministic comparison of localization and office files.
It supports visual diff reporting and multi-version comparison.

Window title in GUI: **Diff View**.

## Key Features

- File vs File comparison with multiple files per side.
- Automatic file pairing by name + manual pairing override.
- Multi-Version chain comparison with a single summary HTML report.
- HTML reports (interactive).
- Excel reports with inline rich diff.
- DOCX Track Changes export for DOCX comparison (when Microsoft Word is available).

## Supported Input Formats

- XLIFF family: `.xliff`, `.xlf`, `.sdlxliff`, `.mqxliff`
- Text: `.txt`, `.srt`
- Office: `.xlsx`, `.xls`, `.pptx`, `.docx`

## GUI Modes

## 1) File vs File

- Drag and drop multiple files to both zones (`File A` / `File B`).
- You can also drag folders into zones. Matching files from the folder are added.
- Zones are clickable; double-click opens file picker.
- Right-click on a tile removes the file from the list.
- File tiles show short names and tooltip with full name.
- Automatic pairing is based on equal filename (case-insensitive).
- Manual pairing:
  - left-click a file in `File A` to select it;
  - files in `File B` become selectable candidates;
  - click a candidate in `File B` to set mapping.
- Tile colors:
  - pale green: matched
  - pale red: unmatched
  - pale orange: currently selected source file
  - dashed border on candidates

## 2) Multi-Version

- Load 2+ files of the same format.
- Reorder by drag-and-drop in the list.
- Comparison is sequential (`v1 -> v2 -> v3 ...`).
- Generates one summary HTML report with filters and per-version highlighting.

## CLI

`main.py` starts GUI with no arguments, and CLI with arguments.

Examples:

```bash
python main.py
python cli.py formats
python cli.py compare tests/fixtures/sample_a.xliff tests/fixtures/sample_b.xliff -o ./output/
python cli.py batch folder_a/ folder_b/ -o ./output/
python cli.py versions v1.xliff v2.xliff v3.xliff -o ./output/
```

Available CLI commands:

- `compare` - compare two files
- `batch` - compare two folders
- `versions` - compare chain of versions
- `formats` - print supported formats

## Report Naming

To avoid accidental overwrites, timestamp now includes seconds where applicable:

- File vs File:
  - `changereport_DD-MM-YY--HH-MM-SS.html`
  - `changereport_DD-MM-YY--HH-MM-SS.xlsx`
- Multi-Version:
  - `versions_summary_DD-MM-YY--HH-MM-SS.html`

## Install From Source

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Build Single EXE (No Console)

Recommended spec:

```bash
pyinstaller main.spec
```

Output:

- `dist\Diffviewer.exe`

The executable uses `Diffviewer.ico` for window/app/taskbar icon.

## Requirements

- Windows
- Python 3.11+
- Microsoft Word (optional, only for DOCX Track Changes export)
