# Diffviewer

> **RU:** Инструмент для визуального сравнения файлов локализации и офисных документов на Windows.
> **EN:** A Windows desktop tool for deterministic visual comparison of localization and office files.

---

## Возможности / Features

| RU | EN |
|----|----|
| Сравнение файл-к-файлу, несколько файлов с каждой стороны | File vs File comparison, multiple files per side |
| Автоматическое сопоставление по имени + ручная настройка | Automatic pairing by filename + manual override |
| Цепочечное сравнение нескольких версий (`v1 → v2 → v3`) | Multi-version chained comparison (`v1 → v2 → v3`) |
| Пакетная обработка папок | Batch folder processing |
| Режим Excel: сравнение колонка-к-колонке с разметкой внутри ячеек | Excel column-by-column mode with inline cell markup |
| Интерактивные HTML-отчёты с фильтрами (All / Changed / Added / Deleted / Modified) | Interactive HTML reports with granular filters |
| Excel-отчёты с встроенным diff | Excel reports with inline rich diff |
| Экспорт DOCX с Track Changes (при наличии Microsoft Word) | DOCX Track Changes export (requires Microsoft Word) |
| Сохранение настроек между запусками | Settings persistence between sessions |
| Поддержка PDF (текстовое извлечение) | PDF support (text extraction) |
| GUI и CLI режимы | GUI and CLI modes |

---

## Поддерживаемые форматы / Supported Formats

| Тип / Type | Форматы / Formats |
|------------|-------------------|
| XLIFF | `.xliff`, `.xlf`, `.sdlxliff`, `.mqxliff` |
| Текст / Text | `.txt`, `.srt` |
| Office | `.xlsx`, `.xls`, `.pptx`, `.docx` |
| PDF | `.pdf` *(требует `pip install pymupdf` / requires `pip install pymupdf`)* |

---

## GUI

### Режим «Файл vs Файл» / File vs File Mode

- Перетащите файлы или папки в зоны **File A** / **File B**
- Двойной клик открывает файловый менеджер
- Правый клик на плашке — удалить файл из списка
- Цвета плашек / Tile colors:
  - 🟢 бледно-зелёный — файл сопоставлен / matched
  - 🔴 бледно-красный — без пары / unmatched
  - 🟠 оранжевый — выбранный источник / selected source

### Режим Excel «Сравнить по колонкам» / Excel Column Comparison

При сравнении двух `.xlsx`-файлов доступен дополнительный режим:

- Сохраняет оригинальную структуру таблицы
- Изменения показываются прямо внутри ячеек
- 🔴 красный зачёркнутый — удалённый текст / deleted text
- 🔵 синий — добавленный текст / added text
- Чёрный — без изменений / unchanged

### Режим «Мультиверсия» / Multi-Version Mode

- Загрузите 2+ файла одного формата
- Переупорядочьте перетаскиванием
- Сравнение идёт последовательно: `v1 → v2 → v3 ...`
- Один сводный HTML-отчёт с фильтрами по версиям

---

## CLI

`main.py` без аргументов запускает GUI. С аргументами — CLI.

```bash
# Запуск GUI / Launch GUI
python main.py

# Сравнить два файла / Compare two files
python cli.py compare old.xliff new.xliff -o ./output/

# Пакетное сравнение папок / Batch compare folders
python cli.py batch folder_a/ folder_b/ -o ./output/

# Сравнение версий / Compare versions chain
python cli.py versions v1.xliff v2.xliff v3.xliff -o ./output/

# Список поддерживаемых форматов / List supported formats
python cli.py formats
```

### Вывод статистики / Statistics output

```
Statistics: total=120 added=5 deleted=3 modified=12 unchanged=100 change%=16.7%
```

---

## Именование отчётов / Report Naming

| Режим / Mode | Файл / File |
|---|---|
| File vs File | `changereport_DD-MM-YY--HH-MM-SS.html` / `.xlsx` |
| Excel column mode | `column_comparison_DD-MM-YY--HH-MM-SS.xlsx` |
| Multi-Version | `versions_summary_DD-MM-YY--HH-MM-SS.html` |

---

## Установка / Installation

### Исполняемый файл / Executable

Скачайте `.exe` из [Releases](../../releases) и запустите — установка не нужна.
Download `.exe` from [Releases](../../releases) and run — no installation needed.

### Из исходников / From Source

```bash
git clone https://github.com/Slipfaith/Diffviewer.git
cd Diffviewer
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

### Сборка EXE / Build EXE

```bash
pyinstaller main.spec
# → dist\Diffviewer.exe
```

---

## Требования / Requirements

- Windows
- Python 3.11+
- Microsoft Word *(опционально / optional — только для DOCX Track Changes export)*

---

## Лицензия / License

MIT
