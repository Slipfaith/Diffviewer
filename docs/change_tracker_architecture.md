# Change Tracker — Архитектурный документ

## 1. Обзор продукта

**Change Tracker** — десктопное Windows-приложение для сравнения файлов перевода и файлов редактора. Программа выявляет добавленные, удалённые и изменённые элементы и генерирует наглядные отчёты.

### Поддерживаемые форматы

| Категория | Форматы |
|-----------|---------|
| XLIFF-семейство | `.xliff`, `.sdlxliff`, `.mqxliff` (MemoQ), `.xlf` |
| Microsoft Office | `.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx` |
| Текстовые | `.txt`, `.srt` |

### Режимы работы

- **Файл vs файл** — сравнение двух файлов одного формата
- **Папка vs папка** — батч-обработка всех файлов с автоматическим сопоставлением по имени
- **Мульти-версия** — сравнение нескольких версий одного файла (цепочка изменений)

### Формат вывода

- **DOCX** → результат в режиме Track Changes (через C# модуль)
- **Все остальные** → HTML-отчёт (интерактивный, в браузере) + Excel с rich text

### Способы запуска

- GUI (Pyside6) с drag & drop и диалогом выбора файлов
- CLI для автоматизации и скриптов

---

## 2. Высокоуровневая архитектура

```
┌─────────────────────────────────────────────────────────┐
│                    ENTRY POINTS                         │
│                                                         │
│   ┌──────────────┐          ┌──────────────────┐        │
│   │  GUI (Pyside6) │          │  CLI (argparse)  │        │
│   │  drag & drop │          │  batch scripts   │        │
│   │  file dialog │          │  CI/CD pipelines │        │
│   └──────┬───────┘          └────────┬─────────┘        │
│          │                           │                  │
│          └───────────┬───────────────┘                  │
│                      ▼                                  │
│          ┌───────────────────────┐                      │
│          │     ORCHESTRATOR      │                      │
│          │                       │                      │
│          │  - определяет формат  │                      │
│          │  - выбирает парсер    │                      │
│          │  - запускает diff     │                      │
│          │  - выбирает генератор │                      │
│          │  - батч-обработка     │                      │
│          │  - мульти-версия      │                      │
│          └───────────┬───────────┘                      │
│                      │                                  │
│    ┌─────────────────┼─────────────────┐                │
│    ▼                 ▼                 ▼                │
│ ┌────────┐    ┌────────────┐    ┌────────────┐         │
│ │PARSERS │    │ DIFF ENGINE│    │  REPORTERS │         │
│ │(плагины)│    │            │    │ (плагины)  │         │
│ └────────┘    └────────────┘    └────────────┘         │
│                                                         │
│          ┌───────────────────────┐                      │
│          │   UNIFIED DATA MODEL  │                      │
│          │  (связывает всё)      │                      │
│          └───────────────────────┘                      │
└─────────────────────────────────────────────────────────┘
                       │
                       ▼ (только для DOCX)
              ┌─────────────────┐
              │  C# .exe модуль │
              │  (Track Changes)│
              └─────────────────┘
```

---

## 3. Unified Data Model (ядро всей системы)

Все форматы приводятся к единой модели **до** сравнения. Это главное архитектурное решение — diff-движок и генераторы отчётов работают только с этой моделью, ничего не зная о форматах файлов.

### 3.1 Segment — единица контента

```
Segment
├── id: str                    # уникальный идентификатор внутри документа
├── source: str | None         # исходный текст (для билингвальных форматов: xliff)
├── target: str                # целевой текст / переведённый текст / контент
├── context: SegmentContext    # метаданные: откуда взят сегмент
├── formatting: list[FormatRun]│ # rich text: жирный, курсив, etc.
└── metadata: dict             # произвольные доп. данные от парсера
```

### 3.2 SegmentContext — откуда взят сегмент

```
SegmentContext
├── file_path: str             # исходный файл
├── location: str              # человекочитаемый путь ("Слайд 3 > Заголовок", "Лист1!B5")
├── position: int              # порядковый номер в документе
└── group: str | None          # логическая группа (trans-unit id, имя листа, номер слайда)
```

### 3.3 FormatRun — фрагмент форматированного текста

```
FormatRun
├── text: str
├── bold: bool
├── italic: bool
├── underline: bool
├── strikethrough: bool
├── color: str | None          # hex цвет
├── font: str | None
└── size: float | None
```

Зачем FormatRun: чтобы в Excel-отчёте с XlsxWriter корректно отображать rich text — каждый FormatRun станет отдельным фрагментом в `write_rich_string()`.

### 3.4 ParsedDocument — результат работы парсера

```
ParsedDocument
├── segments: list[Segment]
├── format_name: str           # "SDLXLIFF", "Excel Workbook", etc.
├── file_path: str
├── metadata: dict             # доп. информация о документе целиком
└── encoding: str | None
```

### 3.5 ChangeRecord — результат работы diff-движка

```
ChangeRecord
├── type: ChangeType           # ADDED | DELETED | MODIFIED | MOVED | UNCHANGED
├── segment_before: Segment | None
├── segment_after: Segment | None
├── text_diff: list[DiffChunk] # пословный/посимвольный diff внутри сегмента
├── similarity: float          # 0.0-1.0, для MODIFIED
└── context: SegmentContext
```

### 3.6 DiffChunk — единица текстового diff

```
DiffChunk
├── type: ChunkType            # EQUAL | INSERT | DELETE
├── text: str
└── formatting: list[FormatRun] | None
```

### 3.7 ComparisonResult — итоговый результат сравнения

```
ComparisonResult
├── file_a: ParsedDocument
├── file_b: ParsedDocument
├── changes: list[ChangeRecord]
├── statistics: ChangeStatistics
└── timestamp: datetime
```

### 3.8 ChangeStatistics

```
ChangeStatistics
├── total_segments: int
├── added: int
├── deleted: int
├── modified: int
├── moved: int
├── unchanged: int
└── change_percentage: float
```

---

## 4. Плагинная система парсеров

### 4.1 Базовый класс

```
BaseParser (ABC)
├── name: str                              # "MemoQ XLIFF Parser"
├── supported_extensions: list[str]        # [".mqxliff"]
├── format_description: str                # для UI
│
├── can_handle(filepath) → bool            # проверка по расширению + валидация
├── parse(filepath) → ParsedDocument       # основная работа
└── validate(filepath) → list[str]         # список ошибок/предупреждений
```

### 4.2 Иерархия парсеров

```
BaseParser (ABC)
│
├── BaseXliffParser                        # общая XLIFF-логика
│   ├── XliffParser                        # стандартный XLIFF 1.2 / 2.0
│   ├── SdlXliffParser                     # SDL Trados SDLXLIFF
│   └── MemoQXliffParser                   # MemoQ XLIFF
│
├── BaseSpreadsheetParser                  # общая логика для таблиц
│   ├── XlsxParser                         # .xlsx через openpyxl
│   └── XlsParser                          # .xls через xlrd
│
├── PptxParser                             # .pptx через python-pptx
│
├── DocxParser                             # .docx через python-docx (только чтение)
│
├── BaseTextParser                         # общая логика для текстовых
│   ├── TxtParser                          # plain text
│   └── SrtParser                          # SubRip subtitles
│
└── (будущие парсеры просто наследуют BaseParser)
```

### 4.3 Автоматическое обнаружение парсеров

Все парсеры лежат в папке `parsers/`. При запуске приложение:

1. Сканирует все `.py` файлы в `parsers/`
2. Импортирует модули
3. Находит все классы-наследники `BaseParser`
4. Регистрирует их в `ParserRegistry`
5. При получении файла — перебирает `can_handle()` и выбирает подходящий

Для добавления нового формата достаточно:
- Создать файл `parsers/my_new_format_parser.py`
- Реализовать класс-наследник `BaseParser`
- Положить файл в папку — всё, он автоматически подхватится

### 4.4 ParserRegistry

```
ParserRegistry
├── _parsers: dict[str, type[BaseParser]]  # extension → parser class
│
├── discover()                             # сканирует папку parsers/
├── get_parser(filepath) → BaseParser      # подбирает парсер по файлу
├── supported_extensions() → list[str]     # для file dialog фильтра
└── register(parser_class)                 # ручная регистрация
```

---

## 5. Diff Engine

### 5.1 Двухуровневое сравнение

**Уровень 1: Сопоставление сегментов** (какой сегмент с каким сравнивать)

```
SegmentMatcher
├── match_by_id(a, b)          # точное совпадение по segment.id (xliff trans-unit)
├── match_by_position(a, b)    # по порядковому номеру (txt, srt)
├── match_by_content(a, b)     # fuzzy matching по тексту (fallback)
└── match(a, b) → MatchResult  # комбинированная стратегия
```

Стратегия выбирается автоматически на основе формата:
- XLIFF-семейство → сначала по ID, потом fuzzy для осиротевших
- Текстовые → по позиции
- Excel → по координатам ячеек (лист + строка + колонка)
- PPTX → по номеру слайда + индексу shape

**Уровень 2: Текстовый diff внутри сегмента** (что именно изменилось)

```
TextDiffer
├── diff_words(a, b) → list[DiffChunk]     # пословное сравнение
├── diff_chars(a, b) → list[DiffChunk]     # посимвольное (для коротких строк)
└── diff_auto(a, b) → list[DiffChunk]      # выбирает стратегию по длине
```

Используются: `difflib.SequenceMatcher` для пословного, `diff-match-patch` для посимвольного.

### 5.2 DiffEngine — фасад

```
DiffEngine
├── compare(doc_a, doc_b) → ComparisonResult
├── compare_multi(docs: list) → list[ComparisonResult]  # цепочка версий
└── _calculate_statistics(changes) → ChangeStatistics
```

### 5.3 Пороги

```
SIMILARITY_THRESHOLD = 0.6    # ниже — считается deleted + added, а не modified
MOVE_DETECTION = True          # искать перемещённые сегменты
FUZZY_MATCH_THRESHOLD = 0.8   # для сопоставления осиротевших сегментов
```

Пороги настраиваемые — через конфиг или UI.

---

## 6. Генераторы отчётов (Reporters)

### 6.1 Базовый класс

```
BaseReporter (ABC)
├── name: str
├── output_extension: str
│
├── generate(result: ComparisonResult, output_path: str) → str
└── supports_rich_text: bool
```

### 6.2 Реализации

```
BaseReporter (ABC)
│
├── HtmlReporter                           # интерактивный HTML
│   ├── Jinja2 шаблон
│   ├── встроенный CSS (self-contained, один файл)
│   ├── inline diff с <ins>/<del>
│   ├── фильтры: Added / Deleted / Modified / All
│   ├── статистика сверху
│   ├── поиск по тексту
│   └── кнопка "Export to Excel" (вызывает ExcelReporter)
│
├── ExcelReporter                          # XlsxWriter с rich text
│   ├── write_rich_string() для inline diff
│   ├── красный + зачёркнутый = удалённое
│   ├── зелёный + подчёркнутый = добавленное
│   ├── автофильтры на колонках
│   ├── закреплённая шапка
│   ├── лист со статистикой
│   └── условное форматирование по типу изменения
│
├── DocxTrackChangesReporter               # обёртка над C# .exe
│   ├── вызов subprocess → docx_compare.exe
│   ├── передача: file_a, file_b, output_path
│   ├── проверка наличия .exe при инициализации
│   └── fallback: ошибка с понятным сообщением
│
└── (будущие: PDF, JSON, CSV — просто новый файл в папке)
```

### 6.3 Автоматический выбор

```
ReporterSelector
├── select(format: str) → list[BaseReporter]
│
│   Логика:
│   - .docx / .doc → [DocxTrackChangesReporter]
│   - всё остальное → [HtmlReporter, ExcelReporter]  # оба сразу
```

---

## 7. Orchestrator

Центральный координатор — единственное место, которое знает обо всех компонентах.

```
Orchestrator
│
├── compare_files(file_a, file_b, output_dir) → list[str]
│   1. ParserRegistry.get_parser(file_a) → parser
│   2. parser.parse(file_a) → doc_a
│   3. parser.parse(file_b) → doc_b
│   4. DiffEngine.compare(doc_a, doc_b) → result
│   5. ReporterSelector.select(format) → reporters
│   6. for reporter in reporters: reporter.generate(result, output)
│   7. return list of output file paths
│
├── compare_folders(folder_a, folder_b, output_dir) → BatchResult
│   1. Сканирует обе папки
│   2. Сопоставляет файлы по имени (с учётом расширений)
│   3. Для каждой пары → compare_files()
│   4. Файлы без пары → помечает как "только в A" / "только в B"
│   5. Генерирует сводный отчёт
│
├── compare_versions(files: list[str], output_dir) → MultiVersionResult
│   1. Сортирует файлы (по имени / дате / пользовательскому порядку)
│   2. Сравнивает цепочкой: [1 vs 2], [2 vs 3], [3 vs 4]...
│   3. Генерирует отчёт с timeline
│
├── on_progress: Callback                  # для прогресс-бара в UI
└── on_error: Callback                     # для обработки ошибок
```

---

## 8. Entry Points

### 8.1 CLI

```
change-tracker compare file_a.xliff file_b.xliff -o ./output/
change-tracker compare --batch folder_a/ folder_b/ -o ./output/
change-tracker compare --versions v1.xliff v2.xliff v3.xliff -o ./output/
change-tracker formats                     # список поддерживаемых форматов
```

Реализация: `argparse` → вызов `Orchestrator`.

### 8.2 GUI (Pyside6)

```
MainWindow
├── FileDropZone (left)                    # drag & drop или кнопка "Browse"
├── FileDropZone (right)                   #
├── ModeSelector                           # File vs File / Folder / Multi-version
├── CompareButton                          # запуск
├── ProgressBar                            #
├── StatusBar                              # текущий файл, статистика
└── ResultPanel                            # ссылки на сгенерированные отчёты
```

GUI вызывает `Orchestrator` в отдельном `QThread`, чтобы не замораживать интерфейс.

---

## 9. C# модуль (docx_compare.exe)

### 9.1 Интерфейс

```
docx_compare.exe <file_a> <file_b> <output> [--author "Change Tracker"]
```

Exit codes: 0 = success, 1 = error (stderr содержит описание).

### 9.2 Реализация

- .NET 8, консольное приложение
- Open XML SDK для генерации Track Changes
- Публикация: `dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true`
- Результат: один .exe файл ~30-60 MB, не требует .NET runtime на машине пользователя

### 9.3 Алгоритм

1. Читает оба DOCX через Open XML SDK
2. Извлекает параграфы с форматированием
3. Сопоставляет параграфы (по позиции + fuzzy)
4. Для изменённых параграфов — генерирует `<w:ins>` / `<w:del>` разметку
5. Сохраняет результат как DOCX с Track Changes

### 9.4 Альтернативный путь

Если C# модуль недоступен, `DocxTrackChangesReporter` может:
- Проверить наличие MS Word → использовать `win32com` для `Document.Compare()`
- Если Word тоже нет → fallback на HTML+Excel отчёт для DOCX (с предупреждением)

---

## 10. Структура проекта

```
change-tracker/
│
├── main.py                                # entry point (GUI или CLI)
├── cli.py                                 # argparse CLI
├── config.py                              # пороги, настройки, пути
│
├── core/
│   ├── __init__.py
│   ├── models.py                          # все dataclasses из раздела 3
│   ├── diff_engine.py                     # DiffEngine, SegmentMatcher, TextDiffer
│   ├── orchestrator.py                    # Orchestrator
│   └── registry.py                        # ParserRegistry, ReporterRegistry
│
├── parsers/
│   ├── __init__.py
│   ├── base.py                            # BaseParser (ABC)
│   ├── xliff_base.py                      # BaseXliffParser
│   ├── xliff_parser.py                    # стандартный XLIFF
│   ├── sdlxliff_parser.py                # SDL Trados
│   ├── memoq_parser.py                   # MemoQ
│   ├── xlsx_parser.py
│   ├── xls_parser.py
│   ├── pptx_parser.py
│   ├── docx_parser.py                     # только чтение
│   ├── txt_parser.py
│   └── srt_parser.py
│
├── reporters/
│   ├── __init__.py
│   ├── base.py                            # BaseReporter (ABC)
│   ├── html_reporter.py                   # Jinja2 + CSS
│   ├── excel_reporter.py                  # XlsxWriter
│   ├── docx_reporter.py                   # обёртка C# .exe
│   └── templates/
│       ├── report.html.j2                 # Jinja2 шаблон
│       └── styles.css                     # встроится в HTML
│
├── ui/
│   ├── __init__.py
│   ├── main_window.py
│   ├── file_drop_zone.py
│   ├── comparison_worker.py               # QThread
│   └── resources/
│       └── icons/
│
├── tools/
│   └── docx_compare.exe                   # скомпилированный C# модуль
│
├── tests/
│   ├── test_models.py
│   ├── test_diff_engine.py
│   ├── test_parsers/
│   │   ├── test_xliff.py
│   │   └── ...
│   ├── test_reporters/
│   └── fixtures/                          # тестовые файлы каждого формата
│
├── requirements.txt
├── setup.py
└── build.spec                             # PyInstaller конфиг
```

---

## 11. Зависимости

### Python (requirements.txt)

```
# UI
Pyside6>=6.5

# XML
lxml>=4.9

# Office форматы (чтение)
openpyxl>=3.1          # .xlsx
xlrd>=2.0              # .xls
python-pptx>=0.6       # .pptx
python-docx>=0.8       # .docx

# Отчёты
XlsxWriter>=3.1        # Excel с rich text
Jinja2>=3.1            # HTML шаблоны

# Diff
diff-match-patch>=20230430  # посимвольный diff

# Упаковка
PyInstaller>=5.0       # сборка .exe
```

### C# (NuGet)

```
DocumentFormat.OpenXml  # Open XML SDK
```

---

## 12. Порядок разработки (дорожная карта)

### Фаза 1: Фундамент
1. `core/models.py` — все dataclasses
2. `core/registry.py` — плагинная система с автодискавери
3. `core/diff_engine.py` — TextDiffer + SegmentMatcher
4. Юнит-тесты для всего выше

### Фаза 2: Первый рабочий пайплайн
5. `parsers/xliff_parser.py` — самый частый формат
6. `parsers/txt_parser.py` — самый простой, для отладки
7. `reporters/html_reporter.py` — наглядная проверка результатов
8. `core/orchestrator.py` — связывает всё вместе
9. `cli.py` — можно запускать из терминала

**Milestone: рабочий CLI для XLIFF + TXT → HTML отчёт**

### Фаза 3: Расширение форматов
10. `parsers/sdlxliff_parser.py`
11. `parsers/memoq_parser.py`
12. `parsers/srt_parser.py`
13. `parsers/xlsx_parser.py` + `xls_parser.py`
14. `parsers/pptx_parser.py`
15. `parsers/docx_parser.py` (только чтение/парсинг)

### Фаза 4: Excel отчёт
16. `reporters/excel_reporter.py` с rich text через XlsxWriter

**Milestone: все форматы + HTML + Excel отчёты**

### Фаза 5: DOCX Track Changes
17. C# модуль `docx_compare.exe`
18. `reporters/docx_reporter.py` (обёртка)

### Фаза 6: Батч и мульти-версия
19. Батч-обработка папок в Orchestrator
20. Мульти-версия (цепочка сравнений)
21. Сводные отчёты для батча

### Фаза 7: GUI
22. Pyside6 интерфейс
23. Drag & drop
24. Прогресс-бар, статистика

### Фаза 8: Упаковка и polish
25. PyInstaller сборка
26. Включение C# .exe в бандл
27. Иконки, about dialog
28. Обработка edge cases, логирование

---

## 13. Ключевые принципы

**Детерминированность.** Никакого AI в ядре. Одинаковые входные файлы всегда дают одинаковый результат.

**Разделение ответственности.** Парсеры не знают про отчёты. Отчёты не знают про форматы. Diff-движок работает только с унифицированной моделью.

**Расширяемость через плагины.** Новый формат = новый файл в папке. Новый тип отчёта = новый файл в папке. Ноль изменений в существующем коде.

**Fail gracefully.** Если парсер не может обработать файл — понятное сообщение об ошибке, а не крэш. Если C# модуль недоступен — fallback на HTML+Excel.

**Тесты на каждом уровне.** Фикстуры реальных файлов каждого формата. Тесты парсеров, diff-движка и генераторов независимо друг от друга.
