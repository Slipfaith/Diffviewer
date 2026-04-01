# Changelog

## [2.4] — 2026-04-01

### Changed

- **Минимальный размер окна** увеличен до 900×620 — зоны drag-and-drop больше не сжимаются до нечитаемого состояния при уменьшении окна.
- **Убраны пороги схожести и нечёткого совпадения из UI** — настройки были избыточны для большинства пользователей; алгоритм использует фиксированные значения (схожесть: 0.60, нечёткое совпадение: 0.80).
- **Обновлена справка** — добавлено подробное описание опции «Не учитывать регистр» с примерами.

---

## [2.3] — 2026-04-01

### Added

- **PDF support** — new parser based on PyMuPDF (`fitz`). Extracts text blocks page-by-page; each block becomes a segment with ID `page{N}_para{M}`. Requires `pip install pymupdf`.
- **Configurable comparison thresholds** — two new spinboxes in the bottom panel:
  - *Порог схожести* (Similarity threshold, default 0.6) — pairs with similarity below this are classified as ADDED/DELETED instead of MODIFIED.
  - *Порог нечёткого совпадения* (Fuzzy match threshold, default 0.8) — minimum similarity for fuzzy segment pairing.
  - Both values persist between sessions via QSettings.
- **Granular HTML report filters** — three new filter buttons alongside "Changed" / "Show All":
  - **Added** (green) — show only added segments.
  - **Deleted** (red) — show only deleted segments.
  - **Modified** (yellow) — show only modified segments.

### Fixed

- **Settings persistence** — output folder and "Не учитывать регистр" checkbox now saved to Windows registry (QSettings) and restored on next launch.
- **Case-insensitive comparison in multi-version and 1-vs-all summaries** — `_build_version_rows` and `_build_one_vs_all_rows` now perform casefold comparison for state classification when `ignore_case=True`. Previously, case-only differences were counted as changes in HTML/Excel summary reports even with the option enabled.
- **Case-insensitive Excel summaries** — `ExcelReporter.generate_versions` and `generate_one_vs_all` now accept and propagate `ignore_case`, so Excel and HTML summaries are consistent.
- **Crash on startup** — fixed `NameError: name 'should_show' is not defined` in `_update_excel_source_controls_visibility`.

---

## [2.2] — earlier

- Ignore-case comparison option added to all modes.
- Multi-pair DOCX track-changes reports per file.
- Excel multi-version "Статус" column for filtering.
- SDLXLIFF strict ID mode; duplicate source deduplication removed.
