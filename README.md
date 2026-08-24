# DocConvert

Convert Excel (`.xlsx`/`.xls`) and Word (`.docx`/`.doc`) documents to **HTML**, **Markdown**, and **JSON** — with full merged-cell support for spreadsheets and a multi-stage cleaning pipeline for Word output.

## Features

- **Excel** — Sheet selection, merged cells (rowspan/colspan), HTML/Markdown/JSON output
- **Word** — `.docx` via `python-docx` + `mammoth`, `.doc` via `textract`
- **Markdown cleaning** — Remove page numbers, duplicate headers, collapse blank lines, normalize whitespace (all configurable)
- **GUI** — Tkinter-based file list, preview, progress bar, overwrite confirmation
- **CLI** — Batch conversion with argparse
- **Async** — Non-blocking background thread; cancel via window close

## Installation

```bash
pip install docconvert
```

For legacy `.xls` support:
```bash
pip install docconvert[xls]
```

For legacy `.doc` support:
```bash
pip install docconvert[doc]
```

All extras:
```bash
pip install docconvert[all]
```

## Quick Start

### GUI

```bash
python main.py
```

### CLI

```bash
# Convert a single file
python main.py convert input.xlsx --format md

# Convert multiple files
python main.py convert file1.xlsx file2.docx --format html -o ./output

# Enhanced Markdown output
python main.py convert input.docx --format md --enhanced

# Select specific Excel sheets
python main.py convert input.xlsx --format json --sheet "Sheet1" --sheet "Sheet2"
```

### Python API

```python
from docconvert.controller import ConversionController

controller = ConversionController()
results = controller.convert_files(
    files=["input.xlsx"],
    output_fmt="html",
    enhanced_md=True,
)

for name, path, error in results:
    if error:
        print(f"Failed: {name} - {error}")
    else:
        print(f"Success: {name} -> {path}")
```

## Supported Formats

| Input       | HTML | Markdown | JSON |
|-------------|------|----------|------|
| `.xlsx`     | ✅   | ✅       | ✅   |
| `.xls`      | ✅   | ✅       | ✅   |
| `.docx`     | ✅   | ✅       | ✅   |
| `.doc`      | ✅   | ✅       | ✅   |

## Configuration

Cleaning rules can be toggled via `AppConfig`:

```python
from docconvert.config import AppConfig

config = AppConfig(
    cleaning_rules={
        "remove_page_numbers": True,
        "remove_duplicate_headers": True,
        "remove_empty_lines": True,
        "normalize_spaces": True,
    }
)
```

## Testing

```bash
pip install -e ".[all]"
python -m unittest discover -s tests -v
```

All 172 tests pass.

## License

MIT
