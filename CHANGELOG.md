# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

<!-- Nothing yet — PR [#2](https://github.com/leop017/DocConvert/pull/2) landed under v2.0.5. -->

***

## [v2.0.5] — 2026-09-02

### Added

* **`docconvert.parsers` subpackage** — full parser layer feeding `Document` / `Element` data models.
  * [`MarkdownParser`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/markdown.py) — headings (H1–H6), fenced code blocks, tables, ordered/unordered lists; strips YAML front-matter when present.
  * [`HtmlParser`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/html.py) — BeautifulSoup-driven extraction of headings, lists, tables, `<pre>` blocks; drops `<script>` and `<style>` nodes automatically.
  * [`PlainTextParser`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/plain.py) — normalizes whitespace and splits long text into paragraphs.
  * [`BaseParser`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/base.py) ABC + [`get_parser()`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/__init__.py) factory.
  * [`models.py`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/parsers/models.py) — `Document`, `Element`, `Chunk` dataclasses shared across parsers and chunkers.

* **`docconvert.chunkers` subpackage** — text chunking utilities for RAG pipelines.
  * [`FixedSizeChunker`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/chunkers/fixed_size.py) — sliding window by character count with configurable overlap.
  * [`SentenceChunker`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/chunkers/sentence.py) — splits on sentence boundaries (`！？。.!`) and respects chunk size.
  * [`MarkdownChunker`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/chunkers/markdown.py) — groups consecutive markdown paragraphs under the same heading anchor.
  * [`BaseChunker`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/chunkers/base.py) ABC + [`get_chunker()`](file:///C:/Users/Administrator/Documents/q/github-codex/docconvert/chunkers/__init__.py) factory.

### Changed

* **CI: bump GitHub Actions** (Dependabot PR [#1](https://github.com/leop017/DocConvert/pull/1), commit `a3d9d63`)
  * `actions/checkout`: `v4` → `v7`
  * `actions/setup-python`: `v5` → `v7`
  * `actions/upload-artifact`: `v4` → `v7`
  * `actions/download-artifact`: `v4` → `v8`
  * `softprops/action-gh-release`: `v2` → `v3`

  All five upgrades were verified by the CI matrix (ubuntu/windows × Python 3.10/3.11) before merge. No source-level changes are required; only the pinned Action versions changed.

### Removed

* `docconvert.chunkers.table_chunker` — was an empty ABC stub with no tests or downstream usage; replaced by the three concrete chunkers above.

### Notes

* `docconvert.parsers.semantic` shim preserved: `BaseParser` is re-exported from there for backwards compat.
* Dependabot is now enabled (`.github/dependabot.yml`) and will open weekly PRs for both the `pip` and `github-actions` ecosystems.

***

## [v2.0.4] — YANKED (2026-09-02)

### Status

* **YANKED** — this tag was published to PyPI and then reverted.

* The release commit `9cc868a` ("bump version to 2.0.4, fix dead PyPI cross-reference link") was reverted by `817467f` ("remove v2.0.4 release, restore Chinese README link") because the only intended change was a documentation tweak, not a feature/bugfix that warranted a version bump.

* The on-disk source remains at **v2.0.3**; no user-visible behavior changed between v2.0.3 and the YANKED v2.0.4.

### Notes

* If you installed `docconvert-local==2.0.4` from PyPI, upgrade to **v2.0.3** (the current published version).

***

## [v2.0.3] — 2026-09-01

### Added

* **Chinese README** ([README\_zh.md](README_zh.md)) — full translation of README.md with links from both sides.

### Changed

* Python API examples updated to use `ConversionController(DEFAULT_CONFIG)` and `output_dir` parameter, matching the [API Reference](docs/api/index.md).

***

## [v2.0.2] — 2026-09-01

### CI / Build

* Fixed Windows CI: textract no longer installed on Windows runners; `TestDocConverterJsonFormat` properly skipped via `import check` guard.

***

## [v2.0.1] — 2026-08-31

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

## [v2.0.0] — 2026-08-30

### Added

* Excel (.xlsx, .xls) and Word (.docx) to Markdown conversion.

* Structured JSON and HTML output modes.

* Cross-platform GUI application.

* CLI tool for headless / pipeline usage.

* PyInstaller builds for Windows, macOS, and Linux.

### Notes

* First public release.

* 100% offline — no API keys, no data leaves your machine.

