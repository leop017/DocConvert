#!/usr/bin/env python3
"""
Regenerate ``doc_convert.py.md`` — the human-readable snapshot of every
``.py`` file under the project root. The committed markdown file is meant
to be checked into the repo so reviewers can read the entire source on
GitHub without cloning; it MUST be regenerated whenever any tracked
``.py`` file changes.

Usage:
    python gen_doc.py            # writes ./doc_convert.py.md
    python gen_doc.py --check    # exits 1 if the file is out of date
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "doc_convert.py.md"

# Order matters: main entry first, then ``app/`` in import-friendly order.
INCLUDE_PATTERNS = [
    "main.py",
    "docconvert/__init__.py",
    "docconvert/chunkers/__init__.py",
    "docconvert/chunkers/table_chunker.py",
    "docconvert/cleaners/__init__.py",
    "docconvert/cleaners/base.py",
    "docconvert/cleaners/word_md.py",
    "docconvert/cli.py",
    "docconvert/config.py",
    "docconvert/controller/__init__.py",
    "docconvert/controller/conversion_controller.py",
    "docconvert/converters/__init__.py",
    "docconvert/converters/base.py",
    "docconvert/converters/doc.py",
    "docconvert/converters/excel.py",
    "docconvert/converters/word.py",
    "docconvert/exporters/__init__.py",
    "docconvert/exporters/base.py",
    "docconvert/exporters/html.py",
    "docconvert/exporters/json_exporter.py",
    "docconvert/exporters/markdown.py",
    "docconvert/gui/__init__.py",
    "docconvert/gui/app.py",
    "docconvert/logger.py",
    "docconvert/models/__init__.py",
    "docconvert/models/models.py",
    "docconvert/parsers/__init__.py",
    "docconvert/parsers/semantic.py",
    "docconvert/utils/__init__.py",
    "docconvert/utils/utils.py",
]


def _slug(rel: str) -> str:
    """Build a GitHub-style heading anchor for ``rel``.

    GitHub lowercases the heading text and replaces every non-alphanumeric
    run with a single ``-``. We mirror that here so the TOC links land on
    the right heading.
    """
    import re
    s = rel.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-")


def build_snapshot() -> str:
    parts: list[str] = []
    total_chars = 0

    parts.append("# DocConvert 全部 Python 源代码快照\n")
    parts.append(
        "> **该文件由脚本自动生成，请勿手动编辑。**\n"
        "> 重新生成命令：`python gen_doc.py`\n"
        "> 校验同步命令：`python gen_doc.py --check`\n"
    )

    parts.append("## 目录\n")
    for rel in INCLUDE_PATTERNS:
        parts.append(f"- [`{rel}`](#{_slug(rel)})\n")
    parts.append("\n---\n\n")

    for rel in INCLUDE_PATTERNS:
        path = ROOT / rel
        if not path.is_file():
            raise FileNotFoundError(f"Tracked file missing: {rel}")
        content = path.read_text(encoding="utf-8")
        total_chars += len(content)
        parts.append(f"## `{rel}`\n\n")
        parts.append("```python\n")
        parts.append(content)
        if not content.endswith("\n"):
            parts.append("\n")
        parts.append("```\n\n")
        parts.append("---\n\n")

    parts.append(
        f"\n<!-- snapshot: {len(INCLUDE_PATTERNS)} files, {total_chars:,} chars "
        f"(regenerate with `python gen_doc.py`) -->\n"
    )
    return "".join(parts)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Exit 1 if doc_convert.py.md is out of date, else 0.",
    )
    args = parser.parse_args()

    snapshot = build_snapshot()

    if args.check:
        if not OUTPUT.is_file():
            print(f"ERROR: {OUTPUT.name} is missing; run gen_doc.py to create it.",
                  file=sys.stderr)
            return 1
        current = OUTPUT.read_text(encoding="utf-8")
        if current == snapshot:
            print(f"{OUTPUT.name} is up to date.")
            return 0
        print(f"ERROR: {OUTPUT.name} is out of date; run gen_doc.py to refresh it.",
              file=sys.stderr)
        return 1

    OUTPUT.write_text(snapshot, encoding="utf-8")
    print(f"Wrote {OUTPUT.name} ({len(snapshot):,} chars, {len(INCLUDE_PATTERNS)} files).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
