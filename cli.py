from __future__ import annotations

import argparse
from pathlib import Path

from core.models import ParseError, UnsupportedFormatError
from core.orchestrator import Orchestrator
from core.registry import ParserRegistry


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="change-tracker")
    subparsers = parser.add_subparsers(dest="command", required=True)

    compare = subparsers.add_parser("compare", help="Compare two files")
    compare.add_argument("file_a", type=str)
    compare.add_argument("file_b", type=str)
    compare.add_argument(
        "-o",
        "--output",
        type=str,
        default="./output/",
        help="Output directory",
    )

    batch = subparsers.add_parser("batch", help="Compare two folders")
    batch.add_argument("folder_a", type=str)
    batch.add_argument("folder_b", type=str)
    batch.add_argument(
        "-o",
        "--output",
        type=str,
        default="./output/",
        help="Output directory",
    )

    versions = subparsers.add_parser("versions", help="Compare multiple versions")
    versions.add_argument("files", nargs="+", type=str)
    versions.add_argument(
        "-o",
        "--output",
        type=str,
        default="./output/",
        help="Output directory",
    )

    subparsers.add_parser("formats", help="List supported formats")
    return parser


def cmd_compare(args: argparse.Namespace) -> int:
    orchestrator = Orchestrator()
    print(f"Comparing: {args.file_a} vs {args.file_b}")
    try:
        outputs = orchestrator.compare_files(args.file_a, args.file_b, args.output)
    except (ParseError, UnsupportedFormatError, RuntimeError) as exc:
        print(f"Error: {exc}")
        return 1
    except Exception as exc:
        print(f"Error: {exc}")
        return 1

    if Path(args.file_a).suffix.lower() == ".docx":
        if not any(Path(output).suffix.lower() == ".docx" for output in outputs):
            print("Microsoft Word not found, generating HTML+Excel reports only")

    result = orchestrator.last_result
    if result is not None:
        stats = result.statistics
        print(
            "Statistics: total={total} added={added} deleted={deleted} "
            "modified={modified} unchanged={unchanged} change%={percent:.1f}%".format(
                total=stats.total_segments,
                added=stats.added,
                deleted=stats.deleted,
                modified=stats.modified,
                unchanged=stats.unchanged,
                percent=stats.change_percentage * 100,
            )
        )
    for output in outputs:
        suffix = Path(output).suffix.lower()
        if suffix == ".html":
            print(f"HTML report: {output}")
        elif suffix == ".xlsx":
            print(f"Excel report: {output}")
        elif suffix == ".docx":
            print(f"Track Changes report: {output}")
        else:
            print(f"Report: {output}")
    return 0


def cmd_batch(args: argparse.Namespace) -> int:
    def on_progress(message: str, value: float) -> None:
        percent = int(value * 100)
        print(f"{message} ({percent}%)")

    orchestrator = Orchestrator(on_progress=on_progress)
    print(f"Batch comparing: {args.folder_a} vs {args.folder_b}")
    try:
        result = orchestrator.compare_folders(args.folder_a, args.folder_b, args.output)
    except Exception as exc:
        print(f"Error: {exc}")
        return 1
    print(
        "Summary: total={total} compared={compared} only_in_a={only_a} only_in_b={only_b} errors={errors}".format(
            total=result.total_files,
            compared=result.compared_files,
            only_a=result.only_in_a,
            only_b=result.only_in_b,
            errors=result.errors,
        )
    )
    if result.summary_report_path:
        print(f"Summary HTML: {result.summary_report_path}")
    if result.summary_excel_path:
        print(f"Summary Excel: {result.summary_excel_path}")
    return 0


def cmd_versions(args: argparse.Namespace) -> int:
    def on_progress(message: str, value: float) -> None:
        percent = int(value * 100)
        print(f"{message} ({percent}%)")

    orchestrator = Orchestrator(on_progress=on_progress)
    print("Comparing versions:")
    for path in args.files:
        print(f"  {path}")
    try:
        result = orchestrator.compare_versions(args.files, args.output)
    except Exception as exc:
        print(f"Error: {exc}")
        return 1
    print(f"Comparisons generated: {len(result.comparisons)}")
    if result.summary_report_path:
        print(f"Summary HTML: {result.summary_report_path}")
    return 0


def cmd_formats() -> int:
    ParserRegistry.discover()
    formats = ParserRegistry.supported_extensions()
    print("Supported formats:")
    for ext in formats:
        print(f"  {ext}")
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.command == "compare":
        return cmd_compare(args)
    if args.command == "batch":
        return cmd_batch(args)
    if args.command == "versions":
        return cmd_versions(args)
    if args.command == "formats":
        return cmd_formats()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
