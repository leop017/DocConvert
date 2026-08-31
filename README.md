# DocConvert

> Convert `.xlsx` / `.xls` / `.docx` / `.doc` → HTML, Markdown, JSON — locally, no cloud upload, no API key.

[![PyPI version](https://img.shields.io/pypi/v/docconvert.svg)](https://pypi.org/project/docconvert/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![Build](https://github.com/leop017/DocConvert/actions/workflows/ci.yml/badge.svg)](https://github.com/leop017/DocConvert/actions)
[![Downloads](https://static.pepy.tech/badge/docconvert/month)](https://pepy.tech/project/docconvert)
[![Stars](https://img.shields.io/github/stars/leop017/DocConvert?style=social)](https://github.com/leop017/DocConvert/stargazers)

## Why DocConvert

| Feature | DocConvert | Online converters | Pandas + python-docx |
|---|:---:|:---:|:---:|
| Converts `.doc` (legacy Word) | ✅ | ⚠️ Partial | ❌ |
| Merged-cell support in Excel | ✅ | ⚠️ Breaks often | ✅ Manual |
| Markdown output with cleaning pipeline | ✅ Configurable | ❌ Raw HTML | ❌ |
| Desktop GUI (no terminal needed) | ✅ | N/A | ❌ |
| Works offline — no data leaves your machine | ✅ | ❌ Upload required | ✅ |
| Batch convert multiple files at once | ✅ | ⚠️ Limited | ❌ |

**Zero configuration. No account. Your files never leave your computer.**

## Features

- **Excel** — Sheet selection, merged cells (rowspan/colspan), HTML / Markdown / JSON
- **Word** — `.docx` via `python-docx` + `mammoth`, legacy `.doc` via `textract`
- **Smart Markdown** — Removes page numbers, duplicate headers, collapses blank lines; all rules configurable
- **GUI** — Tkinter desktop app with file list, preview, progress bar, overwrite protection
- **CLI** — One-line batch conversion via argparse
- **Python API** — Programmatic control with full type hints
- **Executable releases** — Download a standalone `.exe` for Windows / macOS / Linux, no Python install needed

## Installation

```bash
pip install docconvert
```

Optional extras:

```bash
# Legacy .doc support (Linux / macOS only)
pip install docconvert[doc]

# Full feature set including build tools
pip install docconvert[all]
```

## Quick Start

### GUI (interactive)

```bash
python main.py
```

### CLI (batch / scripting)

```bash
# Single file → Markdown
python main.py convert input.xlsx --format md

# Multiple files → HTML into output/
python main.py convert file1.xlsx file2.docx --format html -o ./output

# Enhanced Markdown (cleaning pipeline)
python main.py convert input.docx --format md --enhanced

# Pick specific Excel sheets → JSON
python main.py convert input.xlsx --format json --sheet "Sheet1" --sheet "Sheet2"
```

### Python API

```python
from docconvert.controller import ConversionController

controller = ConversionController()
results = controller.convert_files(
    files=["input.xlsx", "report.docx"],
    output_fmt="html",
    enhanced_md=True,
)

for name, path, error in results:
    if error:
        print(f"Failed: {name} – {error}")
    else:
        print(f"OK: {name} → {path}")
```

## Output Preview

**Input** — an Excel sheet with merged cells:

| Region | Q1 | Q2 |
|:---:|:---:|:---:|
| **North** | 120 | 150 |
| **South** | 90 | 200 |

→ **Markdown output** (auto-cleaned):

```markdown
## Region    Q1    Q2
North       120   150
South        90   200
```

→ **JSON output**:

```json
{
  "Region": ["North", "South"],
  "Q1": [120, 90],
  "Q2": [150, 200]
}
```

## Configuration

Toggle Markdown cleaning rules via `AppConfig`:

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

## Releases (no Python needed)

Standalone executables for Windows, macOS, and Linux are built automatically on each tag push. Download them from [Releases](https://github.com/leop017/DocConvert/releases).

## Project Layout

```
docconvert/
  converters/     # Excel / Word / .doc readers
  cleaners/       # Markdown cleaning pipeline
  exporters/      # HTML / Markdown / JSON output
  controller/     # Orchestration, async, overwrite checks
  gui/            # Tkinter desktop app
  parsers/, chunkers/  # Extension points
tests/
main.py           # GUI / CLI entry point
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT
