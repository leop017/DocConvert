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

For `.doc` support (Linux/macOS only):
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

All tests pass under `unittest discover`.

## Project Layout

```text
docconvert/
  converters/   # excel / word / doc readers
  cleaners/     # Word/Markdown cleaning pipeline
  exporters/    # html / markdown / json exporters
  controller/   # orchestration, async/cancel, overwrite checks
  gui/          # Tkinter application
  parsers/, chunkers/  # reserved extension points
tests/          # unittest suite (mirrors the package layout)
main.py         # GUI / CLI entry point
```

## Maintenance

- Run the full test suite before opening a PR (see Testing above).
- Keep `README.md`, `CONTRIBUTING.md`, and `pyproject.toml` in sync when adding new entry points, extras, or CLI flags.
- Generated sources: regenerate `doc_convert.py.md` via `python gen_doc.py` whenever it changes; otherwise leave it untouched.

## Releases / Executables

End-user releases are produced by GitHub Actions whenever a `v*` tag is pushed (or when the workflow is dispatched manually). Each run builds Windows / macOS / Linux executables via PyInstaller and uploads them to the matching GitHub Release under `https://github.com/leop017/codex/releases`.

To cut a new release:

1. Make sure `main` is green and the version in `pyproject.toml` matches the tag you intend to push.
2. Tag the commit: `git tag -a vX.Y.Z -m "DocConvert X.Y.Z"` and `git push origin vX.Y.Z`.
3. The `release` workflow builds the artefacts and creates the GitHub Release automatically.

If you want to build the executable locally:

```bash
pip install -e ".[all,build]"
python build_scripts/build_exe.py --clean
```

The build script writes an output under `dist/DocConvert-<platform>-<arch>/`.

## License

MIT
