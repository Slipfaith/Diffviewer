from __future__ import annotations

from pathlib import Path
import shutil

import pytest

from core.models import ParseError, UnsupportedFormatError
from core.orchestrator import Orchestrator


FIXTURES = Path(__file__).resolve().parent / "fixtures"


def test_orchestrator_xliff(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(
        str(FIXTURES / "sample_a.xliff"),
        str(FIXTURES / "sample_b.xliff"),
        str(tmp_path),
    )
    assert len(outputs) == 2
    html_file = next(Path(output) for output in outputs if Path(output).suffix == ".html")
    excel_file = next(Path(output) for output in outputs if Path(output).suffix == ".xlsx")
    assert html_file.exists()
    assert excel_file.exists()
    assert html_file.name.startswith("changereport_")
    assert excel_file.name.startswith("changereport_")
    assert html_file.stem == excel_file.stem


def test_orchestrator_txt(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(
        str(FIXTURES / "sample_a.txt"),
        str(FIXTURES / "sample_b.txt"),
        str(tmp_path),
    )
    assert len(outputs) == 2
    html_file = next(Path(output) for output in outputs if Path(output).suffix == ".html")
    excel_file = next(Path(output) for output in outputs if Path(output).suffix == ".xlsx")
    assert html_file.exists()
    assert excel_file.exists()
    assert html_file.name.startswith("changereport_")
    assert excel_file.name.startswith("changereport_")
    assert html_file.stem == excel_file.stem


def test_orchestrator_txt_preserves_literal_entities_in_report(tmp_path: Path) -> None:
    file_a = tmp_path / "a.txt"
    file_b = tmp_path / "b.txt"
    file_a.write_text("Don&#39;t old\n", encoding="utf-8")
    file_b.write_text("Don&#39;t new\n", encoding="utf-8")

    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(str(file_a), str(file_b), str(tmp_path))
    html_file = next(Path(output) for output in outputs if Path(output).suffix == ".html")
    html_content = html_file.read_text(encoding="utf-8")

    assert "Don&amp;#39;t" in html_content
    assert "Don't" not in html_content


def test_orchestrator_non_text_decodes_entities_in_report(tmp_path: Path) -> None:
    file_a = tmp_path / "a.xliff"
    file_b = tmp_path / "b.xliff"
    file_a.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file source-language="en" target-language="en" datatype="plaintext" original="sample.txt">
    <body>
      <trans-unit id="1">
        <source>Don&amp;#39;t</source>
        <target>Old value</target>
      </trans-unit>
    </body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )
    file_b.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file source-language="en" target-language="en" datatype="plaintext" original="sample.txt">
    <body>
      <trans-unit id="1">
        <source>Don&amp;#39;t</source>
        <target>New value</target>
      </trans-unit>
    </body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )

    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(str(file_a), str(file_b), str(tmp_path))
    html_file = next(Path(output) for output in outputs if Path(output).suffix == ".html")
    html_content = html_file.read_text(encoding="utf-8")

    assert "Don't" in html_content
    assert "&#39;" not in html_content
    assert "&amp;#39;" not in html_content


def test_orchestrator_xlsx_with_source_columns(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    outputs = orchestrator.compare_files(
        str(FIXTURES / "sample_a.xlsx"),
        str(FIXTURES / "sample_b.xlsx"),
        str(tmp_path),
        excel_source_column_a="A",
        excel_source_column_b="A",
    )
    assert len(outputs) == 2
    html_file = next(Path(output) for output in outputs if Path(output).suffix == ".html")
    excel_file = next(Path(output) for output in outputs if Path(output).suffix == ".xlsx")
    assert html_file.exists()
    assert excel_file.exists()


def test_orchestrator_invalid_xlsx_source_column(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    with pytest.raises(ParseError):
        orchestrator.compare_files(
            str(FIXTURES / "sample_a.xlsx"),
            str(FIXTURES / "sample_b.xlsx"),
            str(tmp_path),
            excel_source_column_a="A-1",
        )


def test_orchestrator_unsupported_format(tmp_path: Path) -> None:
    file_a = tmp_path / "file_a.unknown"
    file_b = tmp_path / "file_b.unknown"
    file_a.write_text("a", encoding="utf-8")
    file_b.write_text("b", encoding="utf-8")
    orchestrator = Orchestrator()
    with pytest.raises(UnsupportedFormatError):
        orchestrator.compare_files(str(file_a), str(file_b), str(tmp_path))


def test_orchestrator_missing_file(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    missing = tmp_path / "missing.txt"
    other = tmp_path / "other.txt"
    other.write_text("content", encoding="utf-8")
    with pytest.raises(ParseError):
        orchestrator.compare_files(str(missing), str(other), str(tmp_path))


def test_orchestrator_compare_folders(tmp_path: Path) -> None:
    folder_a = tmp_path / "a"
    folder_b = tmp_path / "b"
    folder_a.mkdir()
    folder_b.mkdir()

    shutil.copy(FIXTURES / "sample_a.txt", folder_a / "shared.txt")
    shutil.copy(FIXTURES / "sample_b.txt", folder_b / "shared.txt")

    shutil.copy(FIXTURES / "sample_a.xliff", folder_a / "only_a.xliff")
    shutil.copy(FIXTURES / "sample_b.srt", folder_b / "only_b.srt")

    (folder_a / "bad.unknown").write_text("a", encoding="utf-8")
    (folder_b / "bad.unknown").write_text("b", encoding="utf-8")

    orchestrator = Orchestrator()
    batch = orchestrator.compare_folders(str(folder_a), str(folder_b), str(tmp_path))

    assert batch.total_files == 4
    assert batch.compared_files == 1
    assert batch.only_in_a == 1
    assert batch.only_in_b == 1
    assert batch.errors == 1
    assert batch.summary_report_path is not None
    assert batch.summary_excel_path is not None
    assert Path(batch.summary_report_path).exists()
    assert Path(batch.summary_excel_path).exists()

    statuses = {item.filename: item.status for item in batch.files}
    assert statuses["shared.txt"] == "compared"
    assert statuses["only_a.xliff"] == "only_in_a"
    assert statuses["only_b.srt"] == "only_in_b"
    assert statuses["bad.unknown"] == "error"

    compared = next(item for item in batch.files if item.filename == "shared.txt")
    assert any(path.endswith(".html") for path in compared.report_paths)
    assert any(path.endswith(".xlsx") for path in compared.report_paths)
    errored = next(item for item in batch.files if item.filename == "bad.unknown")
    assert errored.error_message is not None


def test_orchestrator_compare_folders_empty(tmp_path: Path) -> None:
    folder_a = tmp_path / "a"
    folder_b = tmp_path / "b"
    folder_a.mkdir()
    folder_b.mkdir()
    orchestrator = Orchestrator()
    batch = orchestrator.compare_folders(str(folder_a), str(folder_b), str(tmp_path))
    assert batch.total_files == 0
    assert batch.summary_report_path is not None
    assert batch.summary_excel_path is not None
    assert Path(batch.summary_report_path).exists()
    assert Path(batch.summary_excel_path).exists()


def test_orchestrator_compare_file_pairs_single_report(tmp_path: Path) -> None:
    orchestrator = Orchestrator()
    pairs = [
        (str(FIXTURES / "sample_a.txt"), str(FIXTURES / "sample_b.txt")),
        (str(FIXTURES / "sample_a.srt"), str(FIXTURES / "sample_b.srt")),
    ]

    result = orchestrator.compare_file_pairs(pairs, str(tmp_path))
    outputs = [Path(item) for item in result["outputs"]]

    assert len(outputs) == 2
    html_file = next(path for path in outputs if path.suffix == ".html")
    excel_file = next(path for path in outputs if path.suffix == ".xlsx")
    assert html_file.exists()
    assert excel_file.exists()
    assert html_file.name.startswith("changereport_multi_")
    assert excel_file.name.startswith("changereport_multi_")
    assert html_file.stem == excel_file.stem

    file_results = result["file_results"]
    assert len(file_results) == 2
    assert all(item["error"] is None for item in file_results)

    stats = result["statistics"]
    assert stats is not None
    assert stats.total_segments > 0

    html_content = html_file.read_text(encoding="utf-8")
    assert "sample_a.txt vs sample_b.txt" in html_content
    assert "sample_a.srt vs sample_b.srt" in html_content
    assert 'id="file-filter"' in html_content


def test_orchestrator_compare_versions(tmp_path: Path) -> None:
    v1 = tmp_path / "v1.txt"
    v2 = tmp_path / "v2.txt"
    v3 = tmp_path / "v3.txt"
    v1.write_text("line one\n", encoding="utf-8")
    v2.write_text("line one changed\n", encoding="utf-8")
    v3.write_text("line one changed again\n", encoding="utf-8")

    orchestrator = Orchestrator()
    result = orchestrator.compare_versions([str(v1), str(v2), str(v3)], str(tmp_path))
    assert len(result.comparisons) == 2
    assert len(result.documents) == 3
    assert result.report_paths == []
    assert result.comparisons[0].file_a.file_path == str(v1)
    assert result.comparisons[0].file_b.file_path == str(v2)
    assert result.comparisons[1].file_a.file_path == str(v2)
    assert result.comparisons[1].file_b.file_path == str(v3)
    assert result.summary_report_path is not None
    assert Path(result.summary_report_path).exists()
    summary_text = Path(result.summary_report_path).read_text(encoding="utf-8")
    assert "Target: v1.txt" in summary_text
    assert "Target: v2.txt" in summary_text
    assert "Target: v3.txt" in summary_text
    assert "Changes in v1.txt (base)" in summary_text
    assert "Changes in v2.txt" in summary_text
    assert "Changes in v3.txt" in summary_text
    assert "version-ins-1" in summary_text
    assert "version-ins-2" in summary_text
    assert "data-filter=\"all\"" in summary_text
    assert "data-filter=\"changed\"" in summary_text
