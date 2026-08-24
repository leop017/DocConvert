# Contributing

Thanks for your interest in DocConvert!

## Getting Started

1. Fork the repository
2. Clone your fork
3. Install dependencies: `pip install -e ".[all]"`
4. Run tests: `python -m unittest discover -s tests -v`

## Pull Request Guidelines

- Keep changes focused — one feature or fix per PR.
- Add or update tests for any new functionality; the suite uses `unittest discover`.
- Ensure all existing tests pass before submitting.
- Follow the existing code style:
  - Type hints everywhere, plus `from __future__ import annotations` in new modules
  - Prefer pure functions and small classes; touch the largest files (`docconvert/gui/app.py`, `docconvert/controller/conversion_controller.py`) with care
- Write clear commit messages; English or bilingual summaries are both fine.

## Documentation Hygiene

- Update `README.md` and `pyproject.toml` together when you add or remove entry points, extras, or CLI flags.
- If you change Python sources that feed `doc_convert.py.md`, regenerate the snapshot via `python gen_doc.py` (use `python gen_doc.py --check` to verify it is up to date).
- Keep the Issue templates (`.github/ISSUE_TEMPLATE/`) usable; do not rename fields silently.

## Releases

- Releases are driven by the `release` workflow (`.github/workflows/release.yml`). Pushing a `v*` tag (or triggering the workflow manually) builds Windows / macOS / Linux executables via PyInstaller and publishes them as GitHub Release assets.
- Build artefacts are produced by `build_scripts/build_exe.py`; keep that script in sync with `DocConvert.spec` if you change hidden imports or bundled data.
- Do **not** commit the contents of `dist/`. The `dist/` and `build/` directories stay untracked.

## Reporting Issues

Use the GitHub issue tracker (under the project repository). Include:
- DocConvert version (see `pyproject.toml`)
- Python version
- Steps to reproduce
- Expected vs actual behavior
- Input file (if possible)

## Code of Conduct

Be respectful, constructive, and inclusive.
