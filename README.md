# DocConvert

> Clean Markdown & structured data from Excel & Word — built for RAG pipelines and LLM workflows.
> Works 100% offline. No API keys. No data leaves your machine.

[![PyPI version](https://img.shields.io/pypi/v/docconvert-local.svg?cachebust=1725084000)](https://pypi.org/project/docconvert-local/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![Build](https://github.com/leop017/DocConvert/actions/workflows/ci.yml/badge.svg)](https://github.com/leop017/DocConvert/actions)
[![Stars](https://img.shields.io/github/stars/leop017/DocConvert?style=social)](https://github.com/leop017/DocConvert/stargazers)

**中文版**: [README_zh.md](README_zh.md)

## Why DocConvert

Most document-to-Markdown tools convert the file — they don't **clean** it. Raw outputs are full of page numbers, duplicate headers, and whitespace noise that eats your context window and dilutes retrieval quality.

DocConvert was built for people who feed documents into LLMs and need every token to count.

| Feature                              |       DocConvert      | MarkItDown | Pandas + python-docx |
| ------------------------------------ | :-------------------: | :--------: | :------------------: |
| Legacy `.doc` support                |           ✅           |      ❌     |           ❌          |
| Excel merged cells (rowspan/colspan) |           ✅           |     ⚠️     |        Manual        |
| Configurable cleaning pipeline       | ✅ 4 rules, toggle any |      ❌     |           ❌          |
| Batch + specific sheet selection     |           ✅           |      ✅     |           ❌          |
| Desktop GUI (no terminal needed)     |           ✅           |      ❌     |           ❌          |
| PDF / PPT / audio support            |           ❌           |      ✅     |           ❌          |
| MCP server / Claude integration      |           ❌           |      ✅     |           ❌          |
| 100% offline, no cloud dependency    |           ✅           |      ✅     |           ✅          |

**Choose DocConvert if:** you work with Excel/Word documents inside an organization, need legacy `.doc` support, or want a configurable cleaning pipeline before feeding docs into a RAG system.

**Choose MarkItDown if:** you need PDF, PPT, images, or audio conversion, or want MCP/Claude Desktop integration out of the box.

## Use Cases

* **RAG ingestion** — clean Excel financial reports and Word contracts into Markdown ready for embedding

* **LLM context prep** — strip page numbers, duplicates, and noise before chunking

* **Offline compliance** — convert sensitive documents without uploading to any cloud service

* **Batch automation** — convert entire folders of reports into a structured directory

## Installation

```bash
pip install docconvert-local
```

Optional extras:

```bash
# Legacy .doc support (Linux / macOS only)
pip install docconvert-local[doc]

# Full feature set including build tools
pip install docconvert-local[all]
```

## Quick Start

### GUI (interactive)

```bash
python main.py
```

### CLI (batch / scripting)

```bash
# Single file → clean Markdown
python main.py convert input.xlsx --format md

# Batch convert with enhanced cleaning (recommended for RAG)
python main.py convert input.docx --format md --enhanced

# Multiple files → HTML into output/
python main.py convert file1.xlsx file2.docx --format html -o ./output

# Pick specific Excel sheets → JSON
python main.py convert input.xlsx --format json --sheet "Sheet1" --sheet "Sheet2"
```

### Python API

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG

controller = ConversionController(DEFAULT_CONFIG)
results = controller.convert_files(
    files=["input.xlsx", "report.docx"],
    output_fmt="md",
    output_dir="./out",
    enhanced_md=True,
)

for name, path, error in results:
    if error:
        print(f"Failed: {name} – {error}")
    else:
        print(f"OK: {name} → {path}")
```

For advanced usage, see the [API Reference](docs/api/index.md):

* [Module overview](docs/api/index.md#module-overview) ·

* [CLI reference](docs/api/cli.md) ·

* [ConversionController](docs/api/controller.md) ·

* [Converters / Exporters / Cleaners](docs/api/converters.md) ·

* [Config & Models](docs/api/config.md) ·

* [GUI (Tkinter)](docs/api/gui.md)

### RAG Pipeline Integration

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG
from langchain.text_splitter import RecursiveCharacterTextSplitter

controller = ConversionController(DEFAULT_CONFIG)
docs = []
for name, path, error in controller.convert_files(
    files=["contracts/*.docx"],
    output_fmt="md",
    output_dir="./out",
    enhanced_md=True,
):
    if not error:
        with open(path) as f:
            docs.append(f.read())

# Chunk and embed — noise already removed
splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=100)
chunks = splitter.split_text("\n".join(docs))
```

## Output Preview

**Input** — an Excel sheet with merged cells:

|   Region  |  Q1 |  Q2 |
| :-------: | :-: | :-: |
| **North** | 120 | 150 |
| **South** |  90 | 200 |

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

## Smart Cleaning Pipeline

The `--enhanced` flag runs a configurable cleaning pass that removes common document noise before output. Each rule is independently toggleable:

```python
from docconvert.config import AppConfig

config = AppConfig(
    cleaning_rules={
        "remove_page_numbers": True,    # strips 1, 2, 3… and "Page X of Y"
        "remove_duplicate_headers": True,  # deduplicates repeating section titles
        "remove_empty_lines": True,     # collapses excessive blank lines
        "normalize_spaces": True,       # single-spaces text, preserves tables
    }
)
```

All four rules are enabled by default with `--enhanced`. Set any to `False` to keep the raw output.

## Features

* **Excel** — Sheet selection, merged cells (rowspan/colspan), HTML / Markdown / JSON

* **Word** — `.docx` via `python-docx` + `mammoth`, legacy `.doc` via `textract`

* **Smart Markdown** — Removes page numbers, duplicate headers, collapses blank lines; all rules configurable

* **GUI** — Tkinter desktop app with file list, preview, progress bar, overwrite protection

* **CLI** — One-line batch conversion via argparse

* **Python API** — Programmatic control with full type hints

* **Executable releases** — Download a standalone `.exe` for Windows / macOS / Linux, no Python install needed

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
