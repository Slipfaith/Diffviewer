# AGENT.md — Change Tracker

## Что это за проект

Change Tracker — десктопное Windows-приложение на Python (Pyside6) для сравнения файлов перевода и файлов редактора. Программа выявляет добавленные, удалённые и изменённые элементы и генерирует наглядные отчёты.


## Стек

- **Python 3.13+**
- **UI:** PySide6
- **XML:** lxml
- **Office чтение:** openpyxl (xlsx), xlrd (xls), python-pptx, python-docx
- **Excel отчёты:** XlsxWriter (rich text через `write_rich_string`)
- **HTML отчёты:** Jinja2
- **Diff:** difflib (stdlib) + diff-match-patch
- **CLI:** argparse
- **Упаковка:** PyInstaller

## Архитектура

Подробная архитектура описана в `docs/architecture.md`. Ниже — краткая выжимка.

### Принцип работы

```
Файл A ──→ Parser ──→ ParsedDocument (list[Segment])
                                                      ──→ DiffEngine ──→ ComparisonResult ──→ Reporter ──→ отчёт
Файл B ──→ Parser ──→ ParsedDocument (list[Segment])
```

Все форматы приводятся к единой модели данных (Segment) ДО сравнения. Diff-движок и генераторы отчётов работают ТОЛЬКО с унифицированной моделью и ничего не знают о форматах файлов.

### Ключевые компоненты

| Компонент | Путь | Назначение |
|-----------|------|------------|
| Models | `core/models.py` | Все dataclasses: Segment, ChangeRecord, DiffChunk и др. |
| Registry | `core/registry.py` | Автодискавери парсеров и репортеров |
| DiffEngine | `core/diff_engine.py` | Сопоставление сегментов + текстовый diff |
| Orchestrator | `core/orchestrator.py` | Координатор: связывает парсеры, diff, репортеры |
| Parsers | `parsers/*.py` | Плагины-парсеры для каждого формата |
| Reporters | `reporters/*.py` | Плагины-генераторы отчётов |
| UI | `ui/*.py` | Pyside6 интерфейс |
| CLI | `cli.py` | Интерфейс командной строки |

### Структура проекта

```
change-tracker/
├── main.py
├── cli.py
├── config.py
├── core/
│   ├── __init__.py
│   ├── models.py
│   ├── diff_engine.py
│   ├── orchestrator.py
│   └── registry.py
├── parsers/
│   ├── __init__.py
│   ├── base.py
│   ├── xliff_base.py
│   ├── xliff_parser.py
│   ├── sdlxliff_parser.py
│   ├── memoq_parser.py
│   ├── xlsx_parser.py
│   ├── xls_parser.py
│   ├── pptx_parser.py
│   ├── docx_parser.py
│   ├── txt_parser.py
│   └── srt_parser.py
├── reporters/
│   ├── __init__.py
│   ├── base.py
│   ├── html_reporter.py
│   ├── excel_reporter.py
│   ├── docx_reporter.py
│   └── templates/
│       ├── report.html.j2
│       └── styles.css
├── ui/
│   ├── __init__.py
│   ├── main_window.py
│   ├── file_drop_zone.py
│   └── comparison_worker.py
├── tools/
├── tests/
│   ├── test_models.py
│   ├── test_diff_engine.py
│   ├── test_parsers/
│   ├── test_reporters/
│   └── fixtures/
├── docs/
│   └── architecture.md
├── requirements.txt
└── build.spec
```

## Unified Data Model

Это ядро всей системы. Все решения проходят через эту модель.

### Segment

```python
@dataclass
class Segment:
    id: str                           # уникальный ID внутри документа
    source: str | None                # исходный текст (для билингвальных: xliff)
    target: str                       # целевой текст / контент
    context: SegmentContext           # откуда взят сегмент
    formatting: list[FormatRun]       # rich text фрагменты
    metadata: dict                    # произвольные доп. данные
```

### ChangeRecord

```python
@dataclass
class ChangeRecord:
    type: ChangeType                  # ADDED | DELETED | MODIFIED | MOVED | UNCHANGED
    segment_before: Segment | None
    segment_after: Segment | None
    text_diff: list[DiffChunk]        # пословный/посимвольный diff
    similarity: float                 # 0.0-1.0
    context: SegmentContext
```

### FormatRun

```python
@dataclass
class FormatRun:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    color: str | None = None
    font: str | None = None
    size: float | None = None
```

Полные определения всех моделей — в `core/models.py`. При изменении моделей обязательно обнови этот файл и `docs/architecture.md`.

## Правила разработки

### Общие

- **Python 3.11+**, type hints обязательны на всех публичных методах
- **Dataclasses** для всех моделей данных (не dict, не namedtuple)
- **ABC** для базовых классов парсеров и репортеров
- **Никакого AI** в ядре. Diff — детерминированный. Одинаковые входные файлы = одинаковый результат. Всегда.
- **Логирование** через `logging` (не print). Уровень DEBUG для парсеров, INFO для Orchestrator
- Все строки — **UTF-8**
- Конфиг (пороги, пути) — в `config.py`, не хардкодить в модулях

### Плагинная система

- Все парсеры наследуют `BaseParser` из `parsers/base.py`
- Все репортеры наследуют `BaseReporter` из `reporters/base.py`
- Автодискавери: `ParserRegistry` сканирует `parsers/`, находит наследников `BaseParser`, регистрирует
- Для добавления формата — создать файл в `parsers/`, реализовать `BaseParser`. Никаких изменений в другом коде.
- Каждый парсер объявляет `supported_extensions: list[str]` и реализует `can_handle(filepath) → bool`

### Парсеры

- Парсер ОБЯЗАН возвращать `ParsedDocument` с заполненными `Segment.id` и `Segment.context`
- XLIFF-семейство: используй `lxml`, НЕ `xml.etree` (производительность, namespace handling)
- XLIFF парсеры наследуют `BaseXliffParser` — общая логика извлечения trans-unit, различия только в namespace и кастомных атрибутах
- Excel: `openpyxl` для .xlsx, `xlrd` для .xls. Сегмент = ячейка. ID = "Sheet1!B5"
- PPTX: `python-pptx`. Сегмент = текстовый shape. ID = "slide3_shape2"
- DOCX: `python-docx` только для ЧТЕНИЯ. Track Changes генерируется через win32com (Microsoft Word COM)
- SRT: парсить по стандарту SubRip (номер, таймкод, текст). ID = номер субтитра. Таймкод сохранять в metadata
- TXT: одна строка = один сегмент. ID = номер строки

### Diff Engine

- Двухуровневый: сопоставление сегментов (matching) + текстовый diff внутри сегмента
- Matching стратегия зависит от формата:
  - XLIFF → по trans-unit ID, fallback fuzzy
  - TXT/SRT → по позиции
  - Excel → по координатам ячеек
  - PPTX → по номеру слайда + shape index
- Текстовый diff: `difflib.SequenceMatcher` для пословного, `diff-match-patch` для посимвольного
- Порог модификации: `SIMILARITY_THRESHOLD = 0.6` (ниже → deleted + added, не modified)
- Порог fuzzy matching: `FUZZY_MATCH_THRESHOLD = 0.8`
- Пороги берутся из `config.py`

### Репортеры

- **HTML:** self-contained (один файл, CSS встроен), Jinja2 шаблон в `reporters/templates/`
- **Excel:** `XlsxWriter`, rich text через `write_rich_string()`. Удалённое = красный + зачёркнутый. Добавленное = зелёный + подчёркнутый. Автофильтры. Закреплённая шапка. Лист статистики.
- Автовыбор: DOCX → Track Changes, всё остальное → HTML + Excel (оба генерируются)

### UI (Pyside6)

- GUI запускает Orchestrator в `QThread` (никогда в main thread)
- Drag & drop для файлов и папок
- Прогресс через `on_progress` callback из Orchestrator
- File dialog фильтрует по `ParserRegistry.supported_extensions()`

### Тесты

- `pytest` для всего
- Фикстуры: реальные файлы каждого формата в `tests/fixtures/`
- Тестировать парсеры, diff и репортеры НЕЗАВИСИМО друг от друга
- Для парсеров: проверять что `ParsedDocument` содержит ожидаемое количество сегментов с корректными ID
- Для diff: проверять на известных парах (добавление, удаление, модификация, перемещение)
- Для репортеров: проверять что выходной файл создаётся и содержит ожидаемые маркеры

### Обработка ошибок

- Парсер не может прочитать файл → понятное исключение `ParseError(filepath, reason)`
- Неподдерживаемый формат → `UnsupportedFormatError(extension)`
- Microsoft Word не найден → fallback с предупреждением, не крэш
- Кодировка файла неизвестна → попробовать определить через `chardet`, fallback UTF-8

## Режимы сравнения

### Файл vs файл
```
orchestrator.compare_files(file_a, file_b, output_dir)
```
Оба файла должны быть одного формата (одинаковое расширение).

### Папка vs папка (батч)
```
orchestrator.compare_folders(folder_a, folder_b, output_dir)
```
Сопоставление файлов по имени. Файлы без пары помечаются как "только в A" / "только в B". Генерируется сводный отчёт.

### Мульти-версия
```
orchestrator.compare_versions([v1.xliff, v2.xliff, v3.xliff], output_dir)
```
Цепочка сравнений: v1↔v2, v2↔v3. Отчёт с timeline изменений.

## Фазы разработки

При выборе задачи ориентируйся на этот порядок:

1. **core/models.py** — все dataclasses
2. **core/registry.py** — плагинная система
3. **core/diff_engine.py** — matching + text diff
4. **parsers/base.py** + **parsers/xliff_parser.py** + **parsers/txt_parser.py**
5. **reporters/base.py** + **reporters/html_reporter.py**
6. **core/orchestrator.py** + **cli.py**
7. Остальные парсеры (sdlxliff, memoq, srt, xlsx, xls, pptx, docx)
8. **reporters/excel_reporter.py**
9. DOCX Track Changes (win32com) + **reporters/docx_reporter.py**
10. Батч-обработка и мульти-версия в Orchestrator
11. **ui/** — Pyside6 интерфейс
12. PyInstaller сборка

Каждая фаза должна быть покрыта тестами перед переходом к следующей.

## Что НЕ делать

- Не использовать AI/LLM для сравнения текста — только детерминированные алгоритмы
- Не хардкодить поддерживаемые форматы — всё через плагинную систему
- Не использовать `xml.etree.ElementTree` для XLIFF — только `lxml`
- Не использовать `openpyxl` для ЗАПИСИ отчётов — только `XlsxWriter`
- Не блокировать UI thread при сравнении — всегда `QThread`
- Не использовать print — только `logging`
- Не использовать `WidthType.PERCENTAGE` в Open XML — только DXA
