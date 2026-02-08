# Diff View

`Diff View` is a Windows desktop tool for deterministic comparison of localization and office files.
It supports visual diff reporting, multi-version comparison, and QA fix verification workflows.

Window title in GUI: **Diff View**.

## Key Features

- File vs File comparison with multiple files per side.
- Automatic file pairing by name + manual pairing override.
- Multi-Version chain comparison with a single summary HTML report.
- QA Verify mode for TP/FP verification from Excel QA reports.
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

## 3) QA Verify

Purpose: verify whether QA-marked issues from Excel were applied in final translation files.

- Load one or more QA Excel reports (`.xlsx`).
- Load one or more final translation files (`.xliff`, `.xlf`, `.sdlxliff`, `.mqxliff`).
- Column mapping:
  - auto-detection by header heuristics
  - manual override via `Column Mapping`
  - supports multi-sheet reports

Verification logic:

- Mandatory rule: **Original Translation (Excel)** vs **Final Translation (XLIFF)**.
- `Revised Translation` is reference only and does not affect status.
- `TP`:
  - different from final -> `APPLIED`
  - equal to final -> `NOT APPLIED`
  - missing/ambiguous -> `CANNOT VERIFY`
- `FP` -> `NOT APPLICABLE`
- Segment matching priority:
  - Segment ID
  - exact Source
  - normalized Source
  - compact Source

`FileName` support in QA reports:

- If `FileName` column is mapped, row lookup is restricted to that translation file.
- Filename matching is tolerant:
  - path/no path
  - with/without extension stem key
  - copy suffix variants like `(1)`, `(copy)`, `-copy`, `_copy`

QA Verify export:

- Export via `Export Results` button (no output folder selection in this tab).
- `Verification` sheet includes:
  - Source
  - Original Translation
  - Revised Translation
  - Final Translation
  - Expected File
  - Matched File
  - QA Mark
  - Verification Status
  - Matched Segment ID
  - Reason / Comment
  - Report / Sheet / Row
- `Revised Translation` and `Final Translation` are highlighted by diff against `Original Translation`.
- Source/Original/Revised/Final columns use ~250 px width.
- `Summary` sheet includes:
  - status totals
  - per-file counters
  - explicit note for files with no QA issues in report.

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
- QA Verify export default name:
  - `qa_verify_YYYYMMDD_HHMMSS.xlsx`

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

