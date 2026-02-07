from __future__ import annotations

import logging
import os
import time
from pathlib import Path

from core.models import ComparisonResult
from reporters.base import BaseReporter
from reporters.html_reporter import HtmlReporter

try:
    import pythoncom
    import win32com.client
except Exception:
    pythoncom = None
    win32com = None


logger = logging.getLogger(__name__)


class DocxTrackChangesReporter(BaseReporter):
    name = "DOCX Track Changes Reporter"
    output_extension = ".docx"
    supports_rich_text = True

    def __init__(self, author: str = "Change Tracker", startup_timeout: float = 5.0) -> None:
        self.author = author
        self.startup_timeout = startup_timeout

    def is_available(self) -> bool:
        if win32com is None or pythoncom is None:
            return False
        word = None
        try:
            pythoncom.CoInitialize()
            word = self._start_word()
            word.Quit()
            return True
        except Exception:
            return False
        finally:
            if word is not None:
                self._safe_quit(word)
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        if not self.is_available():
            logger.warning(
                "Microsoft Word not found, falling back to HTML report"
            )
            html_path = output_file.with_suffix(".html")
            HtmlReporter().generate(result, str(html_path))
            return str(html_path)

        output_file.parent.mkdir(parents=True, exist_ok=True)
        abs_output = os.path.abspath(str(output_file))
        abs_a = os.path.abspath(result.file_a.file_path)
        abs_b = os.path.abspath(result.file_b.file_path)

        word = None
        doc_a = None
        doc_b = None
        result_doc = None
        try:
            pythoncom.CoInitialize()
            word = self._start_word()
            word.Visible = False
            word.DisplayAlerts = 0

            doc_a = word.Documents.Open(abs_a, ReadOnly=True)
            doc_b = word.Documents.Open(abs_b, ReadOnly=True)

            word.Application.CompareDocuments(
                OriginalDocument=doc_a,
                RevisedDocument=doc_b,
                Destination=2,
                Granularity=1,
                CompareFormatting=True,
                CompareCaseChanges=True,
                CompareWhitespace=True,
                CompareTables=True,
                CompareHeaders=True,
                CompareFootnotes=True,
                CompareTextboxes=True,
                CompareFields=True,
                CompareComments=True,
                RevisedAuthor=self.author,
            )

            result_doc = word.ActiveDocument
            result_doc.SaveAs2(abs_output, FileFormat=12)
        except Exception as exc:
            logger.error("Word automation failed: %s", exc)
            raise RuntimeError(f"Word automation failed: {exc}") from exc
        finally:
            self._safe_close(result_doc)
            self._safe_close(doc_b)
            self._safe_close(doc_a)
            if word is not None:
                self._safe_quit(word)
            if pythoncom is not None:
                pythoncom.CoUninitialize()

        return str(output_file)

    def _start_word(self):
        if win32com is None:
            raise RuntimeError("pywin32 is not installed")

        deadline = time.time() + self.startup_timeout
        last_exc = None
        while time.time() < deadline:
            try:
                word = win32com.client.Dispatch("Word.Application")
                _ = word.Version
                return word
            except Exception as exc:
                last_exc = exc
                time.sleep(0.1)
        raise RuntimeError("Failed to start Microsoft Word") from last_exc

    @staticmethod
    def _safe_close(document) -> None:
        if document is None:
            return
        try:
            document.Close(SaveChanges=0)
        except Exception:
            pass

    @staticmethod
    def _safe_quit(word) -> None:
        try:
            word.Quit()
        except Exception:
            pass
