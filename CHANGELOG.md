# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [v2.0.3] — 2026-09-01

### Added

* **Chinese README** ([README\_zh.md](README_zh.md)) — full translation of README.md with links from both sides.

### Changed

* Python API examples updated to use `ConversionController(DEFAULT_CONFIG)` and `output_dir` parameter, matching the [API Reference](docs/api/index.md).

***

## \[v2.0.2] — 2026-09-01

### CI / Build

* Fixed Windows CI: textract no longer installed on Windows runners; `TestDocConverterJsonFormat` properly skipped via `import check` guard.

***

## \[v2.0.1] — 2026-08-31

### Changed

* **PyPI package renamed** from `docconvert` to `docconvert-local` to resolve a name conflict on PyPI. Update your install command:

  ```bash
  pip install docconvert-local
  ```

* CLI help text and error messages translated to English for broader accessibility.

* `pyproject.toml` description aligned with README narrative.

### Added

* **Automated PyPI publishing**: the release workflow now builds and publishes the package to PyPI in one step.

* `docconvert` CLI entry point registered in `pyproject.toml`.

### Documentation

* README completely rewritten with RAG/LLM pipeline narrative and comparison to MarkItDown.

* Repository metadata and badges updated.

### CI / Build

* Added mypy type-checking to CI; fixed all type errors.

* Stabilized GUI test suite across Windows, macOS, and Linux runners.

* Fixed release workflow YAML syntax and Python inline-script indentation issues.

* Resolved textract optional-dependency install failures on Linux/macOS.

### Fixed

* Repository URLs updated from old `codex` name to `DocConvert`.

* Badge links fixed (PyPI version badge now points to correct package).

***

## \[v2.0.0] — 2026-08-30

### Added

* Excel (.xlsx, .xls) and Word (.docx) to Markdown conversion.

* Structured JSON and HTML output modes.

* Cross-platform GUI application.

* CLI tool for headless / pipeline usage.

* PyInstaller builds for Windows, macOS, and Linux.

### Notes

* First public release.

* 100% offline — no API keys, no data leaves your machine.

