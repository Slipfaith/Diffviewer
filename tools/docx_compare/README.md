# DOCX Compare (Track Changes)

Generates a DOCX with Track Changes markup by comparing two DOCX files.

## Build

```bash
cd tools/docx_compare
dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true -o ../
```

Result: `tools/docx_compare.exe`

## Usage

```bash
tools/docx_compare.exe <file_a> <file_b> <output> [--author "Change Tracker"]
```

Exit codes:
- `0` success
- `1` error (message in stderr)
