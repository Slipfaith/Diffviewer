from __future__ import annotations

import logging
import os
import shutil
import stat
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

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

# Word constants
WD_NO_PROTECTION = -1
WD_FIND_CONTINUE = 1
WD_REPLACE_ALL = 2

COMMON_HTML_ENTITY_REPLACEMENTS: tuple[tuple[str, str], ...] = (
    ("&amp;#39;", "'"),
    ("&#39;", "'"),
    ("&amp;#x27;", "'"),
    ("&#x27;", "'"),
    ("&amp;apos;", "'"),
    ("&apos;", "'"),
)


class DocxTrackChangesReporter(BaseReporter):
    name = "DOCX Track Changes Reporter"
    output_extension = ".docx"
    supports_rich_text = True

    def __init__(self, author: str = "Change Tracker", startup_timeout: float = 5.0) -> None:
        self.author = author
        self.startup_timeout = startup_timeout
        # When a batch session is open, a single Word instance is kept alive
        # across many generate() calls instead of starting/quitting per file.
        self._session_word = None
        self._session_owns_com = False

    def is_available(self) -> bool:
        if win32com is None or pythoncom is None:
            return False
        word = None
        try:
            pythoncom.CoInitialize()
            word = self._start_word()
            return True
        except Exception:
            return False
        finally:
            if word is not None:
                self._safe_quit(word)
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def open_session(self) -> bool:
        """Start a single Word instance to be reused across many generate() calls.

        Returns True if Word is available and a session was started. While a
        session is open, generate() reuses this instance instead of starting and
        quitting Word per file — the dominant cost in batch/folder runs.

        Must be called and closed (close_session) on the same thread, and all
        generate() calls in between must run on that thread too (COM apartment).
        """
        if win32com is None or pythoncom is None:
            return False
        if self._session_word is not None:
            return True
        try:
            pythoncom.CoInitialize()
            self._session_owns_com = True
            word = self._start_word()
            word.Visible = False
            word.DisplayAlerts = 0
            self._session_word = word
            return True
        except Exception:
            self.close_session()
            return False

    def close_session(self) -> None:
        """Quit the reusable Word instance started by open_session()."""
        if self._session_word is not None:
            self._safe_quit(self._session_word)
            self._session_word = None
        if self._session_owns_com:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            self._session_owns_com = False

    def generate(self, result: ComparisonResult, output_path: str) -> str:
        output_file = Path(output_path)
        if output_file.suffix.lower() != self.output_extension:
            output_file = output_file.with_suffix(self.output_extension)

        if win32com is None or pythoncom is None:
            logger.warning(
                "Microsoft Word not found, falling back to HTML report"
            )
            return self._html_fallback(result, output_file)

        # Reuse an already-open batch session instead of cycling Word per file.
        if self._session_word is not None:
            try:
                self._compare_and_save(self._session_word, result, output_file)
                return str(output_file)
            except Exception as exc:
                logger.warning("Word automation failed: %s - falling back to HTML report", exc)
                return self._html_fallback(result, output_file)

        # Standalone path: own Word lifecycle for this single call.
        word = None
        try:
            pythoncom.CoInitialize()
            word = self._start_word()
            word.Visible = False
            word.DisplayAlerts = 0
            self._compare_and_save(word, result, output_file)
            return str(output_file)
        except Exception as exc:
            logger.warning("Word automation failed: %s - falling back to HTML report", exc)
            return self._html_fallback(result, output_file)
        finally:
            if word is not None:
                self._safe_quit(word)
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _compare_and_save(self, word, result: ComparisonResult, output_file: Path) -> None:
        """Run CompareDocuments for one pair on an existing Word instance.

        Opens both source docs, produces the tracked-changes comparison, saves
        it, and always closes the three documents it touched so the Word
        instance is left clean for the next pair.
        """
        output_file.parent.mkdir(parents=True, exist_ok=True)
        abs_output = os.path.abspath(str(output_file))
        abs_a = os.path.abspath(result.file_a.file_path)
        abs_b = os.path.abspath(result.file_b.file_path)

        tmp_dir = None
        doc_a = None
        doc_b = None
        result_doc = None
        try:
            tmp_dir = tempfile.mkdtemp(prefix="diffviewer_docx_")
            work_a = self._prepare_file(abs_a, tmp_dir, "a_")
            work_b = self._prepare_file(abs_b, tmp_dir, "b_")

            doc_a = self._open_document_for_compare(
                word=word,
                file_path=work_a,
                tmp_dir=tmp_dir,
                prefix="a_",
            )

            doc_b = self._open_document_for_compare(
                word=word,
                file_path=work_b,
                tmp_dir=tmp_dir,
                prefix="b_",
            )

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
        finally:
            self._safe_close(result_doc)
            self._safe_close(doc_b)
            self._safe_close(doc_a)
            if tmp_dir is not None:
                shutil.rmtree(tmp_dir, ignore_errors=True)

    @staticmethod
    def _html_fallback(result: ComparisonResult, output_file: Path) -> str:
        html_path = output_file.with_suffix(".html")
        HtmlReporter().generate(result, str(html_path))
        return str(html_path)

    @staticmethod
    def _prepare_file(src_path: str, tmp_dir: str, prefix: str) -> str:
        """Copy file to temp dir and remove flags that force read-only mode."""
        name = prefix + Path(src_path).name
        dst = os.path.join(tmp_dir, name)
        shutil.copy2(src_path, dst)

        # Remove read-only attribute
        st = os.stat(dst)
        os.chmod(dst, st.st_mode | stat.S_IWRITE | stat.S_IREAD)

        # Remove Mark of the Web (Zone.Identifier ADS) - causes Protected View
        try:
            zone_path = dst + ":Zone.Identifier"
            if os.path.exists(zone_path):
                os.remove(zone_path)
        except OSError:
            pass
        try:
            os.remove(dst + ":Zone.Identifier")
        except OSError:
            pass

        DocxTrackChangesReporter._strip_docx_protection_flags(dst)

        return dst

    def _open_document_for_compare(
        self,
        *,
        word,
        file_path: str,
        tmp_dir: str,
        prefix: str,
    ):
        doc = self._open_word_document(word, file_path)
        self._normalize_document_for_editing(doc)
        if not self._is_document_read_only(doc):
            self._decode_common_html_entities_in_document(doc)
            return doc

        editable_copy = os.path.join(tmp_dir, f"{prefix}editable_{Path(file_path).name}")
        doc.SaveAs2(editable_copy, FileFormat=12)
        self._safe_close(doc)

        reopened = self._open_word_document(word, editable_copy)
        self._normalize_document_for_editing(reopened)
        if self._is_document_read_only(reopened):
            raise RuntimeError(f"Unable to open editable copy for comparison: {file_path}")
        self._decode_common_html_entities_in_document(reopened)
        return reopened

    @staticmethod
    def _open_word_document(word, file_path: str):
        kwargs = {
            "ReadOnly": False,
            "AddToRecentFiles": False,
            "Revert": False,
            "ConfirmConversions": False,
            "OpenAndRepair": True,
            "NoEncodingDialog": True,
        }
        try:
            return word.Documents.Open(file_path, **kwargs)
        except TypeError:
            return word.Documents.Open(file_path, ReadOnly=False, AddToRecentFiles=False)

    def _normalize_document_for_editing(self, doc) -> None:
        self._unprotect_document(doc)
        for attr_name, value in (("Final", False), ("ReadOnlyRecommended", False)):
            try:
                setattr(doc, attr_name, value)
            except Exception:
                continue

    @staticmethod
    def _is_document_read_only(doc) -> bool:
        try:
            return bool(getattr(doc, "ReadOnly"))
        except Exception:
            return False

    @staticmethod
    def _strip_docx_protection_flags(docx_path: str) -> None:
        settings_member = "word/settings.xml"
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        try:
            with zipfile.ZipFile(docx_path, "r") as archive:
                if settings_member not in archive.namelist():
                    return
                settings_xml = archive.read(settings_member)
        except Exception:
            return

        try:
            root = ET.fromstring(settings_xml)
        except ET.ParseError:
            return

        changed = False
        for local_name in ("documentProtection", "readOnlyRecommended"):
            tag_name = f"{{{word_ns}}}{local_name}"
            for element in list(root.findall(tag_name)):
                root.remove(element)
                changed = True

        if not changed:
            return

        patched_settings = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        temp_archive_path = f"{docx_path}.tmp"

        try:
            with zipfile.ZipFile(docx_path, "r") as source_archive:
                with zipfile.ZipFile(temp_archive_path, "w") as target_archive:
                    for info in source_archive.infolist():
                        payload = (
                            patched_settings
                            if info.filename == settings_member
                            else source_archive.read(info.filename)
                        )
                        target_archive.writestr(info, payload)
            os.replace(temp_archive_path, docx_path)
        finally:
            if os.path.exists(temp_archive_path):
                try:
                    os.remove(temp_archive_path)
                except OSError:
                    pass

    def _decode_common_html_entities_in_document(self, doc) -> None:
        for find_text, replace_text in COMMON_HTML_ENTITY_REPLACEMENTS:
            self._replace_all_in_document(doc, find_text=find_text, replace_text=replace_text)

    def _replace_all_in_document(self, doc, *, find_text: str, replace_text: str) -> None:
        try:
            story_range = doc.StoryRanges
        except Exception:
            story_range = None

        if story_range is None:
            content = getattr(doc, "Content", None)
            if content is not None:
                self._replace_all_in_range(content, find_text=find_text, replace_text=replace_text)
            return

        visited_ranges: set[int] = set()
        current_range = story_range
        while current_range is not None and id(current_range) not in visited_ranges:
            visited_ranges.add(id(current_range))
            self._replace_all_in_range(
                current_range,
                find_text=find_text,
                replace_text=replace_text,
            )
            try:
                current_range = current_range.NextStoryRange
            except Exception:
                break

    @staticmethod
    def _replace_all_in_range(range_obj, *, find_text: str, replace_text: str) -> None:
        try:
            find = range_obj.Find
        except Exception:
            return

        try:
            find.ClearFormatting()
        except Exception:
            pass
        try:
            find.Replacement.ClearFormatting()
        except Exception:
            pass

        for attr_name, attr_value in (
            ("Text", find_text),
            ("Forward", True),
            ("Wrap", WD_FIND_CONTINUE),
            ("Format", False),
            ("MatchCase", False),
            ("MatchWholeWord", False),
            ("MatchWildcards", False),
            ("MatchSoundsLike", False),
            ("MatchAllWordForms", False),
        ):
            try:
                setattr(find, attr_name, attr_value)
            except Exception:
                continue
        try:
            find.Replacement.Text = replace_text
        except Exception:
            pass

        try:
            find.Execute(Replace=WD_REPLACE_ALL)
        except TypeError:
            try:
                find.Execute(find_text, False, False, False, False, False, True, WD_FIND_CONTINUE, False, replace_text, WD_REPLACE_ALL)
            except Exception:
                pass
        except Exception:
            pass

    @staticmethod
    def _unprotect_document(doc) -> None:
        """Remove document editing restrictions if present."""
        try:
            if doc.ProtectionType != WD_NO_PROTECTION:
                doc.Unprotect("")
        except Exception:
            try:
                if doc.ProtectionType != WD_NO_PROTECTION:
                    doc.Unprotect()
            except Exception:
                pass

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
