# DocConvert 全部 Python 源代码快照
> **该文件由脚本自动生成，请勿手动编辑。**
> 重新生成命令：`python gen_doc.py`
> 校验同步命令：`python gen_doc.py --check`
## 目录
- [`main.py`](#main-py)
- [`docconvert/__init__.py`](#docconvert-init-py)
- [`docconvert/chunkers/__init__.py`](#docconvert-chunkers-init-py)
- [`docconvert/chunkers/table_chunker.py`](#docconvert-chunkers-table-chunker-py)
- [`docconvert/cleaners/__init__.py`](#docconvert-cleaners-init-py)
- [`docconvert/cleaners/base.py`](#docconvert-cleaners-base-py)
- [`docconvert/cleaners/word_md.py`](#docconvert-cleaners-word-md-py)
- [`docconvert/cli.py`](#docconvert-cli-py)
- [`docconvert/config.py`](#docconvert-config-py)
- [`docconvert/controller/__init__.py`](#docconvert-controller-init-py)
- [`docconvert/controller/conversion_controller.py`](#docconvert-controller-conversion-controller-py)
- [`docconvert/converters/__init__.py`](#docconvert-converters-init-py)
- [`docconvert/converters/base.py`](#docconvert-converters-base-py)
- [`docconvert/converters/doc.py`](#docconvert-converters-doc-py)
- [`docconvert/converters/excel.py`](#docconvert-converters-excel-py)
- [`docconvert/converters/word.py`](#docconvert-converters-word-py)
- [`docconvert/exporters/__init__.py`](#docconvert-exporters-init-py)
- [`docconvert/exporters/base.py`](#docconvert-exporters-base-py)
- [`docconvert/exporters/html.py`](#docconvert-exporters-html-py)
- [`docconvert/exporters/json_exporter.py`](#docconvert-exporters-json-exporter-py)
- [`docconvert/exporters/markdown.py`](#docconvert-exporters-markdown-py)
- [`docconvert/gui/__init__.py`](#docconvert-gui-init-py)
- [`docconvert/gui/app.py`](#docconvert-gui-app-py)
- [`docconvert/logger.py`](#docconvert-logger-py)
- [`docconvert/models/__init__.py`](#docconvert-models-init-py)
- [`docconvert/models/models.py`](#docconvert-models-models-py)
- [`docconvert/parsers/__init__.py`](#docconvert-parsers-init-py)
- [`docconvert/parsers/semantic.py`](#docconvert-parsers-semantic-py)
- [`docconvert/utils/__init__.py`](#docconvert-utils-init-py)
- [`docconvert/utils/utils.py`](#docconvert-utils-utils-py)

---

## `main.py`

```python
#!/usr/bin/env python3
"""
文档转换工具 - Entry Point

Usage:
    python main.py                  # Launch GUI
    python main.py convert <file>   # CLI conversion mode
"""

from __future__ import annotations

import sys

from docconvert.logger import setup_logging


def main():
    setup_logging()
    if len(sys.argv) > 1 and sys.argv[1] == 'convert':
        from docconvert.cli import main_cli
        sys.exit(main_cli(sys.argv[1:]))
    else:
        import tkinter as tk
        from docconvert.gui import DocConvertApp
        root = tk.Tk()
        DocConvertApp(root)
        root.mainloop()


if __name__ == '__main__':
    main()
```

---

## `docconvert/__init__.py`

```python
"""DocConvert - 文档转换工具

The GUI is imported lazily so CLI / library use (``python main.py convert``
or ``from docconvert.controller import ConversionController``) does not
require the Tkinter package.
"""

__all__ = ["DocConvertApp"]


def __getattr__(name: str):
    if name == "DocConvertApp":
        from docconvert.gui.app import DocConvertApp
        return DocConvertApp
    raise AttributeError(f"module 'docconvert' has no attribute {name!r}")
```

---

## `docconvert/chunkers/__init__.py`

```python
from docconvert.chunkers.table_chunker import BaseChunker

__all__ = [
    "BaseChunker",
]
```

---

## `docconvert/chunkers/table_chunker.py`

```python
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class BaseChunker(ABC):
    @abstractmethod
    def chunk(self, content: Any, **kwargs) -> list[Any]:
        ...
```

---

## `docconvert/cleaners/__init__.py`

```python
from docconvert.cleaners.base import BaseCleaner
from docconvert.cleaners.word_md import WordMdCleaner

__all__ = [
    "BaseCleaner",
    "WordMdCleaner",
]
```

---

## `docconvert/cleaners/base.py`

```python
from __future__ import annotations

from abc import ABC, abstractmethod


class BaseCleaner(ABC):

    @abstractmethod
    def clean(self, content: str, **kwargs) -> str:
        ...
```

---

## `docconvert/cleaners/word_md.py`

```python
from __future__ import annotations

import re
from typing import Optional

from docconvert.cleaners.base import BaseCleaner
from docconvert.config import AppConfig, DEFAULT_CONFIG


_UNESCAPE_PATTERN = re.compile(r'\\([\\\[\]\-.\(\)])')
_HORIZ_WS_PATTERN = re.compile(r'[ \t]+')
_FULLWIDTH_SPACE = '\u3000'
# A line that opens or closes a fenced code block (``` or ~~~), possibly
# with a language tag: "```python". Used to keep code-block whitespace intact.
_FENCE_PATTERN = re.compile(r'^\s*(?:`{3,}|~{3,})')


class WordMdCleaner(BaseCleaner):
    """
    Defense-in-depth post-processor for Word -> Markdown output.

    WordConverter already strips header/footer references in the DOCX XML
    before conversion. This cleaner is the regex safety net for whatever
    page-number artifacts leak through mammoth, plus general whitespace
    cleanup that mammoth cannot infer from Word's XML.

    Page-number rules (line-level, anchored to ``strip()`` so they only
    match standalone markers, never page numbers embedded in body text):

    - Inline refs: [1], [12]
    - Chinese: 第3页， 第 5 页
    - Bordered: - 1 -, -12-
    - Western: Page 1, page 5, Pág. 3, P. 7
    - Multi-form: Page 1 of 10, Page 3 / 10

    Document-level rules applied after the line-level pass:

    - ``remove_duplicate_headers`` — drop consecutive duplicate lines
      (the typical leak when mammoth flattens page headers / footers
      into the body of every page). Only consecutive duplicates are
      removed; non-adjacent repeats are kept.
    - ``remove_empty_lines`` — collapse runs of ≥2 empty lines down to
      a single empty line. Whitespace-only lines count as empty.
    - ``normalize_spaces`` — convert U+3000 (full-width space) and tabs
      to a single space, collapse internal runs of horizontal whitespace,
      and strip trailing whitespace. Leading whitespace is preserved so
      indented markdown (code blocks, list items, tables) is not
      destroyed.

    All four rules skip lines inside fenced (````` ``` ```` / ``~~~``) or
    indented (≥4 leading spaces) code blocks: page markers, blank lines,
    repeated lines and internal spacing are all significant inside code,
    and the cleaner must not corrupt it.

    Rule activation is driven by ``AppConfig.cleaning_rules``. The four
    documented keys all have working implementations now:

    - ``remove_page_numbers`` (default True) → enables all 5 page rules
    - ``remove_duplicate_headers`` (default True)
    - ``remove_empty_lines`` (default True)
    - ``normalize_spaces`` (default True)

    When constructed without a config, ``DEFAULT_CONFIG`` is used and
    every rule is active (backward-compatible default).
    """

    _RULES = (
        re.compile(r'^\[\d+\]$'),
        re.compile(r'^第\s*\d+\s*页$'),
        re.compile(r'^-\s*\d+\s*-$'),
        re.compile(r'^page\s*\d+\s*(?:of|/)\s*\d+\s*$', re.IGNORECASE),
        re.compile(r'^(?:Page|Pág\.|P\.)\s*\d+\s*$', re.IGNORECASE),
    )

    _LINE_RULE_GROUPS: dict[str, tuple[int, ...]] = {
        "remove_page_numbers": (0, 1, 2, 3, 4),
    }

    def __init__(self, config: Optional[AppConfig] = None):
        cfg = config or DEFAULT_CONFIG
        rules = cfg.cleaning_rules
        self._active_line_rules = self._resolve_line_rules(rules)
        self._collapse_empty = bool(rules.get("remove_empty_lines", False))
        self._normalize_spaces = bool(rules.get("normalize_spaces", False))
        self._dedupe_consecutive = bool(rules.get("remove_duplicate_headers", False))

    def _resolve_line_rules(self, cleaning_rules: dict) -> tuple:
        active: list[int] = []
        for key, indices in self._LINE_RULE_GROUPS.items():
            if cleaning_rules.get(key, False):
                active.extend(indices)
        return tuple(self._RULES[i] for i in active)

    def clean(self, content: str, **kwargs) -> str:
        if not any((
            self._active_line_rules,
            self._collapse_empty,
            self._normalize_spaces,
            self._dedupe_consecutive,
        )):
            return content

        lines = content.split('\n')
        # Classify code-block lines ONCE so every rule below can leave code
        # untouched: page markers, blank lines, repeated lines and internal
        # spacing are all significant inside code.
        code_mask = self._code_mask(lines)

        # Stage 1: line-level rules (page numbers). Lines matching a
        # rule are dropped; everything else keeps its original whitespace.
        if self._active_line_rules:
            lines, code_mask = self._apply_line_rules(lines, code_mask)

        # Stage 2: collapse empty lines. Run before normalize so that
        # runs of whitespace-only lines (which normalize would turn into
        # empty strings) are deduped exactly once.
        if self._collapse_empty:
            lines, code_mask = self._collapse_empty_lines(lines, code_mask)

        # Stage 3: normalize spaces within each line. Run before dedupe
        # so that lines differing only in whitespace are treated as
        # duplicates.
        if self._normalize_spaces:
            lines = self._normalize_lines_spaces(lines, code_mask)

        # Stage 4: drop consecutive duplicate lines.
        if self._dedupe_consecutive:
            lines = self._dedupe_consecutive_lines(lines, code_mask)

        return '\n'.join(lines)

    def _code_mask(self, lines: list[str]) -> list[bool]:
        """Boolean per line: True when the line lives inside a code block.

        Fenced blocks (````` ``` ```` / ``~~~``) are tracked by toggling on
        their delimiter lines; indented blocks are any line with ≥4 leading
        spaces/tabs. Delimiter lines themselves count as code so they are
        never altered or removed by a later rule.
        """
        in_fence = False
        mask: list[bool] = []
        for line in lines:
            if _FENCE_PATTERN.match(line):
                in_fence = not in_fence
                mask.append(True)
                continue
            mask.append(in_fence or self._is_code_line(line))
        return mask

    def _apply_line_rules(self, lines: list[str], code_mask: list[bool]):
        kept: list[str] = []
        kept_mask: list[bool] = []
        for line, code in zip(lines, code_mask):
            stripped = line.strip()
            if not stripped:
                kept.append(line)
                kept_mask.append(code)
                continue
            if code:
                # Page-number markers inside code are content (e.g. a
                # ``[1]`` array index or a ``Page 1`` literal); never strip.
                kept.append(line)
                kept_mask.append(True)
                continue
            # Normalize mammoth's backslash-escaped markdown punctuation
            # (e.g. ``\[1\]`` → ``[1]``, ``\- 1 \-`` → ``- 1 -``,
            # ``Pág\. 3`` → ``Pág. 3``) so the page-number rules match
            # the real shape of the marker rather than its escaped form.
            normalized = _UNESCAPE_PATTERN.sub(r'\1', stripped)
            if any(rx.match(normalized) for rx in self._active_line_rules):
                continue
            kept.append(line)
            kept_mask.append(code)
        return kept, kept_mask

    @staticmethod
    def _collapse_empty_lines(lines: list[str], code_mask: list[bool]):
        result: list[str] = []
        result_mask: list[bool] = []
        prev_empty = False
        prev_code = False
        for line, code in zip(lines, code_mask):
            is_empty = not line.strip()
            # Blank lines inside a code block are significant and are not
            # collapsed; blank runs outside code still become one line.
            if is_empty and prev_empty and not (code and prev_code):
                continue
            # Normalize the kept empty line to a bare ``""`` so the
            # output is canonical regardless of whether the input used
            # ``"   "`` or ``"\t"`` as its whitespace-only filler.
            result.append("" if is_empty else line)
            result_mask.append(code)
            prev_empty = is_empty
            prev_code = code
        return result, result_mask

    def _normalize_lines_spaces(self, lines: list[str], code_mask: list[bool]) -> list[str]:
        result: list[str] = []
        for line, code in zip(lines, code_mask):
            result.append(self._normalize_line_spaces(line, code_line=code))
        return result

    @staticmethod
    def _is_code_line(line: str) -> bool:
        """True when ``line`` starts an indented markdown code block.

        A line indented by ≥4 spaces/tabs is treated as code by CommonMark,
        where internal spacing is significant. Non-code lines fall through
        to the ordinary whitespace normalization.
        """
        return len(line) - len(line.lstrip(' \t')) >= 4

    @staticmethod
    def _normalize_line_spaces(line: str, code_line: bool = False) -> str:
        # Convert full-width space (U+3000) used in CJK typography to a
        # regular space so downstream consumers can tokenize uniformly.
        line = line.replace(_FULLWIDTH_SPACE, ' ')
        if not line.strip():
            return line
        if code_line:
            # Inside a fenced or indented code block, internal spacing is
            # meaningful — collapsing it would corrupt the code. Only strip
            # trailing whitespace.
            return line.rstrip()
        # Preserve leading whitespace (indentation matters in markdown
        # for code blocks, list items, tables) while collapsing internal
        # runs of horizontal whitespace and stripping trailing whitespace.
        match = re.match(r'^([ \t]*)(.*?)([ \t]*)$', line, re.DOTALL)
        if not match:
            return line.rstrip()
        leading, middle, _trailing = match.groups()
        middle = _HORIZ_WS_PATTERN.sub(' ', middle)
        return leading + middle

    @staticmethod
    def _dedupe_consecutive_lines(lines: list[str], code_mask: list[bool]) -> list[str]:
        # Only collapse runs of identical non-empty lines. Empty lines
        # were already handled by ``_collapse_empty_lines``. Two
        # non-adjacent identical lines are kept as-is, and repeated lines
        # inside code blocks are legitimate and kept too.
        result: list[str] = []
        prev: Optional[str] = None
        prev_code = False
        for line, code in zip(lines, code_mask):
            if line.strip() and line == prev and not (code and prev_code):
                continue
            result.append(line)
            prev = line
            prev_code = code
        return result
```

---

## `docconvert/cli.py`

```python
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docconvert.config import DEFAULT_CONFIG
from docconvert.controller import ConversionController
from docconvert.logger import setup_logging
from docconvert.models import ProgressEvent


def _cli_progress(event: ProgressEvent):
    if event.message:
        print(f'\r[{int(event.progress * 100):3d}%] {event.message:<50s}', file=sys.stderr, end='')
    if event.done:
        print(file=sys.stderr)


class _HelpFormatter(argparse.HelpFormatter):
    """Preserves newlines in description, epilog, and all help strings."""

    def _fill_text(self, text, width, indent):
        if text:
            return ''.join(indent + line + '\n' for line in text.splitlines())
        return ''

    def _split_lines(self, text, width):
        return text.splitlines() if text else []


def main_cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description='DocConvert - 文档转换工具 (CLI)',
        formatter_class=_HelpFormatter,
        epilog=(
            '示例:\n'
            '  python main.py convert input.xlsx --format md\n'
            '  python main.py convert input.docx --format html -o ./output\n'
            '  python main.py convert file1.xlsx file2.xlsx --format json\n'
        ),
    )

    subparsers = parser.add_subparsers(dest='command', help='子命令')

    convert_parser = subparsers.add_parser('convert', help='转换文件')
    convert_parser.add_argument('files', nargs='+', help='输入文件路径')
    convert_parser.add_argument(
        '--format', '-f',
        choices=['html', 'md', 'json'],
        default='html',
        help='输出格式 （默认： html)',
    )
    convert_parser.add_argument(
        '--output', '-o',
        default=None,
        help='输出目录 （默认： 输入文件所在目录）',
    )
    convert_parser.add_argument(
        '--enhanced', '-e',
        action='store_true',
        help='启用增强 Markdown 输出',
    )
    convert_parser.add_argument(
        '--sheet', '-s',
        action='append',
        help='指定 Excel 工作表 （可重复使用， 默认： 全部）',
    )
    convert_parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='详细日志输出',
    )

    args = parser.parse_args(argv)

    if args.command != 'convert':
        parser.print_help()
        return 1

    setup_logging(level='DEBUG' if getattr(args, 'verbose', False) else 'INFO')

    controller = ConversionController(DEFAULT_CONFIG)

    all_results: list[tuple[str, str, str | None]] = []
    failed = 0

    for filepath in args.files:
        p = Path(filepath)
        if not p.exists():
            print(f'错误： 文件不存在 - {filepath}', file=sys.stderr)
            all_results.append((p.name, '', '文件不存在'))
            failed += 1
            continue

        ext = p.suffix.lower()
        if ext not in {'.xlsx', '.xls', '.docx', '.doc'}:
            print(f'错误： 不支持的文件格式 - {filepath}', file=sys.stderr)
            all_results.append((p.name, '', '不支持的文件格式'))
            failed += 1
            continue

        try:
            convert_results = controller.convert_files(
                files=[filepath],
                output_fmt=args.format,
                output_dir=args.output,
                enhanced_md=args.enhanced,
                sheets=args.sheet,
                progress_callback=_cli_progress,
            )
            for name, path, err in convert_results:
                all_results.append((name, path, err))
                if err:
                    failed += 1
                    print(f'失败： {name} - {err}', file=sys.stderr)
                else:
                    print(f'成功： {name} -> {path}')
        except Exception as e:
            print(f'错误： {filepath} - {e}', file=sys.stderr)
            all_results.append((p.name, '', str(e)))
            failed += 1

    if failed:
        print(f'\n处理完成： {len(all_results)} 个文件， {failed} 个失败')
        return 1

    print(f'\n处理完成： {len(all_results)} 个文件， 全部成功')
    return 0


if __name__ == '__main__':
    sys.exit(main_cli())
```

---

## `docconvert/config.py`

```python
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class AppConfig:
    chunk_size: int = 1000
    max_rows: int = 0
    markdown_style: str = "github"
    preview_chars: int = 5000
    preview_lines: int = 150
    large_file_size: int = 20 * 1024 * 1024

    cleaning_rules: dict[str, bool] = field(default_factory=lambda: {
        "remove_page_numbers": True,
        "remove_duplicate_headers": True,
        "remove_empty_lines": True,
        "normalize_spaces": True,
    })


DEFAULT_CONFIG = AppConfig()
```

---

## `docconvert/controller/__init__.py`

```python
from docconvert.controller.conversion_controller import ConversionController

__all__ = [
    "ConversionController",
]
```

---

## `docconvert/controller/conversion_controller.py`

```python
from __future__ import annotations

import os
from pathlib import Path
from threading import Thread, Event, Lock
from typing import Callable, Optional

from docconvert.config import AppConfig, DEFAULT_CONFIG
from docconvert.converters import ExcelConverter, WordConverter, DocConverter
from docconvert.logger import get_logger
from docconvert.models import ProgressEvent
from docconvert.utils import clean_filename, get_excel_sheet_names

ProgressCallback = Callable[[ProgressEvent], None]


class ConversionController:

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG
        self.logger = get_logger()
        self._cancel_event = Event()
        self._done_event = Event()
        self._thread: Optional[Thread] = None
        self._running = False
        self._lock = Lock()
        self.last_results: list[tuple[str, str, Optional[str]]] = []
        self.was_cancelled: bool = False
        self.last_error: Optional[str] = None

    @property
    def is_running(self) -> bool:
        return self._running

    def cancel(self):
        self._cancel_event.set()

    def reset_cancel(self):
        self._cancel_event.clear()

    def set_config(self, config: AppConfig):
        """Replace the active config. Should only be called when idle."""
        if self._running:
            self.logger.warning("无法在转换进行中更新 config,已忽略")
            return False
        self.config = config
        return True

    @staticmethod
    def _batch_parent(files: list[str], output_dir: Optional[str] = None) -> Path:
        """Return the single directory a batch converts into.

        Every output of a batch goes to one directory: the explicit
        ``output_dir`` or, when omitted, the first input file's parent.
        ``convert_files`` and ``check_overwrite_paths`` must agree on this
        value or the overwrite confirmation will not match reality, so both
        resolve it through this helper.
        """
        if output_dir:
            return Path(output_dir)
        return Path(files[0]).resolve().parent if files else Path.cwd()

    @staticmethod
    def _make_unique_stem(base_stem: str, source_path: str, used: set[str]) -> str:
        """Return a batch-unique output stem.

        All files in a batch are written to a single directory (either the
        explicit ``output_dir`` or the first input file's parent), so two
        input files that share a basename would produce the same output
        path and the second would silently overwrite the first. Disambiguate
        subsequent collisions with the source's parent-directory name, then a
        numeric suffix as a last resort. ``used`` is mutated in place.
        """
        if base_stem not in used:
            used.add(base_stem)
            return base_stem
        parent_name = clean_filename(Path(source_path).resolve().parent.name)
        suffix = parent_name if parent_name and parent_name != base_stem else ''
        candidate = f'{base_stem}_{suffix}' if suffix else base_stem
        n = 2
        while candidate in used:
            candidate = f'{base_stem}_{suffix}_{n}' if suffix else f'{base_stem}_{n}'
            n += 1
        used.add(candidate)
        return candidate

    def _get_converter(self, ext: str):
        if ext in ('.xlsx', '.xls'):
            c = ExcelConverter(self.config)
        elif ext == '.docx':
            c = WordConverter(self.config)
        elif ext == '.doc':
            c = DocConverter(self.config)
        else:
            raise ValueError(f'不支持的文件格式： {ext}')
        c.cancel_check = self._cancel_event.is_set
        return c

    def convert_files(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        enhanced_md: bool = False,
        sheets: Optional[list[str]] = None,
        progress_callback: Optional[ProgressCallback] = None,
    ) -> list[tuple[str, str, Optional[str]]]:
        """
        Convert files synchronously.
        Returns list of (filename, output_path, error_or_None).

        Sets ``self.was_cancelled = True`` if the loop was aborted because the
        cancel event was set mid-batch. The caller is expected to surface this
        to the user (e.g., show a partial-results dialog instead of a green
        "completed" badge).
        """
        self.reset_cancel()
        self.was_cancelled = False
        results: list[tuple[str, str, Optional[str]]] = []

        if not files:
            return results

        parent = self._batch_parent(files, output_dir)
        if output_dir and not parent.is_dir():
            try:
                parent.mkdir(parents=True, exist_ok=True)
                self.logger.info("已创建输出目录： %s", parent)
            except OSError as e:
                raise ValueError(
                    f'无法创建输出目录： {parent} - {e}'
                ) from e

        total = len(files)
        used_stems: set[str] = set()
        for idx, input_path in enumerate(files):
            if self._cancel_event.is_set():
                self.was_cancelled = True
                self.logger.info("转换已取消")
                break

            if not os.path.exists(input_path):
                err = '文件不存在'
                fname_missing = Path(input_path).name
                results.append((fname_missing, '', err))
                self._report(progress_callback, ProgressEvent(
                    message=f'文件不存在： {fname_missing}',
                    error=err,
                    progress=(idx + 1) / total,
                ))
                continue

            fname = Path(input_path).name
            self._report(progress_callback, ProgressEvent(
                message=f'处理文件 {idx + 1}/{total}: {fname}',
                progress=(idx + 0.5) / total,
            ))

            ext = Path(input_path).suffix.lower()
            base_stem = clean_filename(Path(input_path).stem)
            unique_stem = self._make_unique_stem(base_stem, input_path, used_stems)
            try:
                converter = self._get_converter(ext)
                # ``sheets`` is Excel-only; passing it to Word/Doc converters
                # raises TypeError (they explicitly reject it), which would
                # otherwise fail every Word/Doc file in a mixed batch.
                convert_kwargs = {
                    'enhanced_md': enhanced_md,
                    'stem_override': unique_stem,
                }
                if ext in ('.xlsx', '.xls'):
                    convert_kwargs['sheets'] = sheets
                file_results, file_errors = converter.convert(
                    input_path, output_fmt, parent, **convert_kwargs
                )
                for name, path in file_results:
                    results.append((name, path, None))
                    self.logger.info("转换成功： %s -> %s", name, path)
                for name, err in file_errors:
                    results.append((name, '', err))
                    self.logger.warning("转换失败 [%s]: %s", name, err)
                    self._report(progress_callback, ProgressEvent(
                        message=f'转换失败： {name}',
                        error=err,
                        progress=(idx + 1) / total,
                    ))
            except Exception as e:
                self.logger.error("转换异常 [%s]: %s", fname, str(e))
                results.append((fname, '', str(e)))
                self._report(progress_callback, ProgressEvent(
                    message=f'转换异常： {fname}',
                    error=str(e),
                    progress=(idx + 1) / total,
                ))

            self._report(progress_callback, ProgressEvent(
                message=f'完成 {idx + 1}/{total}',
                progress=(idx + 1) / total,
            ))
            if self._cancel_event.is_set():
                self.was_cancelled = True

        if self.was_cancelled:
            self._report(progress_callback, ProgressEvent(
                message='转换已取消',
                done=True,
            ))
        else:
            self._report(progress_callback, ProgressEvent(
                message='转换完成',
                progress=1.0,
                done=True,
            ))
        return results

    def convert_files_async(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        enhanced_md: bool = False,
        sheets: Optional[list[str]] = None,
        progress_callback: Optional[ProgressCallback] = None,
    ):
        """Start conversion in background thread."""
        with self._lock:
            if self._running:
                self.logger.warning("转换任务已在运行中")
                return False
            self._running = True
            self._cancel_event.clear()
            self._done_event.clear()
            self.was_cancelled = False
            self.last_error = None

        def _run():
            try:
                self.last_results = self.convert_files(
                    files=files,
                    output_fmt=output_fmt,
                    output_dir=output_dir,
                    enhanced_md=enhanced_md,
                    sheets=sheets,
                    progress_callback=progress_callback,
                )
            except Exception as e:
                self.logger.error("转换任务异常： %s", str(e))
                self.last_results = []
                self.last_error = str(e)
            finally:
                with self._lock:
                    self._running = False
                self._done_event.set()

        self._thread = Thread(target=_run, daemon=True)
        self._thread.start()
        return True

    def wait_for_completion(self, timeout: Optional[float] = None) -> bool:
        """Wait for the background thread to complete. Returns True if completed."""
        if self._thread is None:
            return True
        self._thread.join(timeout=timeout)
        return not self._running

    def wait_done(self, timeout: Optional[float] = None) -> bool:
        """Event-based wait for completion. Returns True if done within timeout."""
        if self._thread is None:
            return True
        return self._done_event.wait(timeout=timeout)

    def _report(
        self,
        callback: Optional[ProgressCallback],
        event: ProgressEvent,
    ):
        if callback:
            try:
                callback(event)
            except Exception as e:
                self.logger.debug("Progress callback error: %s", e)

    def check_overwrite_paths(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        sheets: Optional[list[str]] = None,
    ) -> list[str]:
        """Pre-compute output paths and return existing ones.

        Uses the same batch directory resolution as ``convert_files``
        (``_batch_parent``), so the reported paths are exactly the ones a
        conversion run would write — including when no ``output_dir`` is
        given and a multi-directory batch all lands in the first file's
        parent.
        """
        existing: list[str] = []
        used_stems: set[str] = set()
        parent = self._batch_parent(files, output_dir)
        for input_path in files:
            if not os.path.exists(input_path):
                continue
            base_stem = clean_filename(Path(input_path).stem)
            unique_stem = self._make_unique_stem(base_stem, input_path, used_stems)
            try:
                paths = self._compute_output_paths(
                    input_path, output_fmt, str(parent), sheets,
                    stem_override=unique_stem,
                )
            except Exception as e:
                self.logger.warning(
                    "无法计算输出路径 [%s]: %s （转换时将报错）", input_path, e
                )
                continue
            for p in paths:
                if os.path.exists(p):
                    existing.append(p)
        return existing

    def _compute_output_paths(
        self,
        input_path: str,
        output_fmt: str,
        output_dir: Optional[str] = None,
        sheets: Optional[list[str]] = None,
        stem_override: Optional[str] = None,
    ) -> list[str]:
        """Return the list of output paths that ``input_path`` would produce.

        Order:
        - .docx → [stem_doc.<fmt>]
        - .doc  → [stem.<fmt>]
        - .xlsx / .xls → [stem_<sheet>.<fmt>] per requested sheet, or
          [stem_<sheet>.<fmt>] for every sheet in the workbook when
          ``sheets`` is None.
        - other → [] (unsupported extension)
        """
        from docconvert.utils import clean_filename

        if not os.path.exists(input_path):
            return []

        if output_dir:
            parent = Path(output_dir)
        else:
            parent = Path(input_path).parent

        stem = Path(input_path).stem
        stem_clean = stem_override or clean_filename(stem)
        ext = Path(input_path).suffix.lower()

        if ext == '.docx':
            return [str(parent / f'{stem_clean}_{clean_filename("doc")}.{output_fmt}')]
        if ext == '.doc':
            return [str(parent / f'{stem_clean}.{output_fmt}')]
        if ext in ('.xlsx', '.xls'):
            if not sheets:
                sheets = get_excel_sheet_names(input_path, ext)
            # Mirror ExcelConverter's per-sheet suffix dedup so the
            # reported paths are the ones a conversion run actually writes.
            from docconvert.utils import unique_cleaned_suffixes
            clean_names = unique_cleaned_suffixes(sheets)
            return [
                str(parent / f'{stem_clean}_{cn}.{output_fmt}')
                for cn in clean_names
            ]
        return []
```

---

## `docconvert/converters/__init__.py`

```python
from docconvert.converters.base import BaseConverter
from docconvert.converters.excel import ExcelConverter
from docconvert.converters.word import WordConverter
from docconvert.converters.doc import DocConverter

__all__ = [
    "BaseConverter",
    "ExcelConverter",
    "WordConverter",
    "DocConverter",
]
```

---

## `docconvert/converters/base.py`

```python
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, Optional

from docconvert.config import AppConfig, DEFAULT_CONFIG
from docconvert.logger import get_logger


class BaseConverter(ABC):

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG
        self.logger = get_logger()
        self.cancel_check: Optional[Callable[[], bool]] = None

    @abstractmethod
    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        ...

    def _export(self, content: str, output_fmt: str) -> str:
        from docconvert.exporters import get_exporter
        exporter = get_exporter(output_fmt)
        return exporter.export(content)
```

---

## `docconvert/converters/doc.py`

```python
from __future__ import annotations

import html as html_mod
from pathlib import Path
from typing import Optional

from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter
from docconvert.utils import clean_filename, decode_text, html_to_md


class DocConverter(BaseConverter):

    def __init__(self, config: Optional[AppConfig] = None):
        super().__init__(config)

    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        enhanced_md = kwargs.get('enhanced_md', False)
        stem_override = kwargs.get('stem_override')
        if kwargs.get('sheets') is not None:
            raise TypeError('DocConverter does not accept "sheets" parameter')
        results: list[tuple[str, str]] = []
        errors: list[tuple[str, str]] = []
        if self.cancel_check and self.cancel_check():
            return results, errors
        try:
            name, path = self._convert_doc(
                input_path, output_fmt, parent, enhanced_md, stem_override
            )
            results.append((name, path))
        except Exception as e:
            self.logger.error("Doc转换失败： %s", str(e))
            errors.append((Path(input_path).name, str(e)))
        return results, errors

    def _convert_doc(
        self, input_path: str, output_fmt: str, parent: Path,
        enhanced_md: bool = False, stem_override: Optional[str] = None,
    ) -> tuple[str, str]:
        import textract

        self.logger.info("提取文本： %s", input_path)
        raw_bytes = textract.process(input_path)
        text = decode_text(raw_bytes)
        stem = stem_override or clean_filename(Path(input_path).stem)
        source_name = Path(input_path).name

        if output_fmt == 'html':
            content = (
                f'<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n'
                f'    <meta charset="UTF-8">\n'
                f'    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
                f'    <title>{html_mod.escape(Path(input_path).stem)}</title>\n'
                f'    <style>\n'
                f'        body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; line-height: 1.6; }}\n'
                f'        pre {{ white-space: pre-wrap; word-wrap: break-word; }}\n'
                f'    </style>\n</head>\n<body>\n<pre>\n{html_mod.escape(text)}\n</pre>\n</body>\n</html>'
            )
            output_name = f'{stem}.html'
        elif output_fmt == 'md':
            if enhanced_md:
                wrapped = f'<pre>\n{html_mod.escape(text)}\n</pre>'
                content = html_to_md(wrapped)
            else:
                content = text
            content = f'<!-- source: {source_name} | format: doc -->\n\n{content}'
            output_name = f'{stem}.md'
        elif output_fmt == 'json':
            content = {'source': source_name, 'content': text}
            output_name = f'{stem}.json'
        else:
            raise ValueError(f'不支持的输出格式： {output_fmt}')

        output_path = str(parent / output_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self._export(content, output_fmt))
        return output_name, output_path
```

---

## `docconvert/converters/excel.py`

```python
from __future__ import annotations

import html as html_mod
import json
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook

from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter
from docconvert.models import MergeInfo
from docconvert.utils import (
    clean_filename,
    escape_md_cell,
    html_to_md,
    safe_str,
    unique_cleaned_suffixes,
)

class _XlsMergeRange:
    """Mimics openpyxl's merged cell range for use with _build_merged_map."""
    def __init__(self, min_col: int, min_row: int, max_col: int, max_row: int):
        self.bounds = (min_col, min_row, max_col, max_row)


class ExcelConverter(BaseConverter):

    def __init__(self, config: Optional[AppConfig] = None):
        super().__init__(config)

    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        results: list[tuple[str, str]] = []
        errors: list[tuple[str, str]] = []

        ext = Path(input_path).suffix.lower()
        engine = 'xlrd' if ext == '.xls' else 'openpyxl'

        if ext == '.xls':
            try:
                import xlrd
            except ImportError:
                raise RuntimeError('需安装 xlrd 库以处理 .xls 文件')

        sheets_param = kwargs.get('sheets')
        enhanced_md = kwargs.get('enhanced_md', False)
        stem_override = kwargs.get('stem_override')

        if sheets_param is not None:
            all_sheet_names = self._get_all_sheet_names(input_path, ext)
            existing = [sn for sn in sheets_param if sn in all_sheet_names]
            missing = [sn for sn in sheets_param if sn not in all_sheet_names]
            for sn in missing:
                errors.append((sn, '工作表不存在'))
            if not existing:
                return results, errors
            all_sheets = pd.read_excel(
                input_path, sheet_name=existing, engine=engine
            )
            if not isinstance(all_sheets, dict):
                all_sheets = {existing[0]: all_sheets}
            # Only load merged-cell metadata for sheets we will actually
            # convert — avoids the (small but real) overhead of opening
            # the workbook a second time for unrequested sheets.
            sheet_names = existing
            sheets_to_convert = existing
        else:
            all_sheets = pd.read_excel(input_path, sheet_name=None, engine=engine)
            sheet_names = list(all_sheets.keys())
            sheets_to_convert = sheet_names

        merged_cache = self._load_merged_cache(input_path, ext, sheet_names)

        # Sheet names that clean to the same value (e.g. ``Q"1`` and
        # ``Q<1`` → ``Q_1``) must not share an output path, or the later
        # sheet would silently overwrite the earlier one. Resolve unique
        # suffixes up front so the emitted filenames are unambiguous.
        sn_overrides = dict(
            zip(sheets_to_convert, unique_cleaned_suffixes(sheets_to_convert))
        )

        for sn in sheets_to_convert:
            if self.cancel_check and self.cancel_check():
                self.logger.info("转换已取消")
                break
            try:
                df = all_sheets.get(sn)
                if df is None:
                    errors.append((sn, '工作表不存在'))
                    continue
                mr = merged_cache.get(sn) if merged_cache else None
                name, path = self._convert_sheet(
                    input_path, df, sn, output_fmt, parent, mr,
                    enhanced_md=enhanced_md, stem_override=stem_override,
                    sn_override=sn_overrides.get(sn),
                )
                results.append((name, path))
            except Exception as e:
                self.logger.error("工作表转换失败 [%s]: %s", sn, str(e))
                errors.append((sn, str(e)))

        return results, errors

    def _load_merged_cache(
        self, input_path: str, ext: str, sheet_names: list[str]
    ) -> Optional[dict[str, list]]:
        if ext == '.xls':
            return self._load_merged_cache_xls(input_path, sheet_names)
        if ext != '.xlsx':
            return None
        try:
            wb = load_workbook(input_path, data_only=True)
        except Exception as e:
            self.logger.warning("打开工作簿失败： %s", e)
            return None
        try:
            cache: dict[str, list] = {}
            for sn in sheet_names:
                try:
                    cache[sn] = list(wb[sn].merged_cells.ranges)
                except (KeyError, AttributeError) as e:
                    self.logger.debug("工作表 '%s' 合并单元格读取失败： %s", sn, e)
                    cache[sn] = []
            return cache
        except Exception as e:
            self.logger.warning("合并单元格读取失败： %s", e)
            return None
        finally:
            wb.close()

    @staticmethod
    def _get_all_sheet_names(input_path: str, ext: str) -> list[str]:
        from docconvert.utils import get_excel_sheet_names
        return get_excel_sheet_names(input_path, ext)

    @staticmethod
    def _load_merged_cache_xls(
        input_path: str, sheet_names: list[str]
    ) -> dict[str, list]:
        import xlrd
        wb = xlrd.open_workbook(input_path, formatting_info=True)
        try:
            cache: dict[str, list] = {}
            for sn in sheet_names:
                try:
                    ws = wb.sheet_by_name(sn)
                except (KeyError, xlrd.XLRDError):
                    cache[sn] = []
                    continue
                merged = []
                for rlo, rhi, clo, chi in ws.merged_cells:
                    merged.append(_XlsMergeRange(clo + 1, rlo + 1, chi, rhi))
                cache[sn] = merged
            return cache
        finally:
            wb.release_resources()

    def _build_merged_map(self, merged_ranges: list) -> dict:
        merged_map = {}
        for merged in merged_ranges:
            min_col, min_row, max_col, max_row = merged.bounds
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    key = (row, col)
                    if key not in merged_map:
                        merged_map[key] = MergeInfo(
                            rowspan=max_row - min_row + 1,
                            colspan=max_col - min_col + 1,
                            is_master=(row == min_row and col == min_col),
                            is_merged=True,
                            min_row=min_row,
                            min_col=min_col,
                            max_row=max_row,
                            max_col=max_col,
                        )
                    else:
                        merged_map[key].is_master = False
        return merged_map

    def _df_to_rows(self, df: pd.DataFrame) -> list[list[str]]:
        rows_data = []
        rows_data.append([safe_str(c) for c in df.columns.tolist()])
        for _, row in df.iterrows():
            rows_data.append([safe_str(v) for v in row.tolist()])
        return rows_data

    @staticmethod
    def _filter_merges_for_dropped_rows(
        merged_ranges: list, dropped_df_indices: set
    ) -> list:
        """Filter merge ranges against the rows that ``dropna`` removed.

        ``merged_ranges`` use 1-based workbook row numbers; ``pd.read_excel``
        consumes the first row as the header, so workbook row N (N >= 2)
        maps to DataFrame index (N - 2). The header row (workbook row 1)
        has no DataFrame index (its formula ``row - 2`` yields ``-1``) and
        is never dropped.

        A merge is kept only when none of its rows landed in
        ``dropped_df_indices``. If any leg was dropped, the surviving rows
        below the gap shift upward and no longer align with the merge's
        original row coordinates, so keeping the merge would mask a real
        cell (silent data loss) or visually jump over the gap.

        This uniformly covers header-anchored merges (``min_row == 1``):
        their master lives in the consumed-header region, but their
        vertical legs can still collide with a dropped interior row — e.g.
        ``A1:A3`` with an empty row 3 removed by ``dropna`` would otherwise
        make the master's rowspan swallow the independent ``A4`` cell that
        shifted up into its place.
        """
        filtered: list = []
        for merged in merged_ranges:
            min_col, min_row, max_col, max_row = merged.bounds
            spans_dropped = any(
                (row - 2) in dropped_df_indices
                for row in range(min_row, max_row + 1)
            )
            if not spans_dropped:
                filtered.append(merged)
        return filtered

    @staticmethod
    def _remap_merged_rows(merged_ranges: list, df: "pd.DataFrame") -> list:
        """Translate workbook row numbers to post-``dropna`` rendered rows.

        ``merged_ranges`` arrive in 1-based workbook coordinates. After
        ``dropna`` removes fully-empty rows, the surviving data rows are
        re-indexed sequentially (workbook row R maps to rendered row
        ``2 + position``). Build that mapping and re-base every merge's
        row bounds onto the rendered coordinates so the merge lands on the
        correct cells.

        A merge that survives filtering cannot span a dropped row, so all
        of its rows exist in the mapping; if one does not (defensive),
        drop the merge rather than render it misaligned.
        """
        wb_to_rendered = {1: 1}
        for pos, orig_idx in enumerate(df.index):
            wb_to_rendered[orig_idx + 2] = 2 + pos
        remapped: list = []
        for merged in merged_ranges:
            min_col, min_row, max_col, max_row = merged.bounds
            if min_row not in wb_to_rendered or max_row not in wb_to_rendered:
                continue
            remapped.append(
                _XlsMergeRange(
                    min_col,
                    wb_to_rendered[min_row],
                    max_col,
                    wb_to_rendered[max_row],
                )
            )
        return remapped

    @staticmethod
    def _iter_merge_legs(info):
        """Yield every (row, col) covered by info.

        Used by the merge guard in _convert_sheet to walk every cell
        a merge occupies, including its slave legs.
        """
        for row in range(info.min_row, info.max_row + 1):
            for col in range(info.min_col, info.max_col + 1):
                yield (row, col)

    def _convert_sheet(
        self,
        input_path: str,
        df: pd.DataFrame,
        sheet_name: str,
        output_fmt: str,
        parent: Path,
        merged_ranges: Optional[list] = None,
        enhanced_md: bool = False,
        stem_override: Optional[str] = None,
        sn_override: Optional[str] = None,
    ) -> tuple[str, str]:
        if df.empty:
            raise ValueError('工作表为空')

        all_nan_mask = df.isna().all(axis=1)
        dropped_df_indices = set(df.index[all_nan_mask].tolist())
        if merged_ranges and dropped_df_indices:
            merged_ranges = self._filter_merges_for_dropped_rows(
                merged_ranges, dropped_df_indices
            )

        df = df.dropna(how='all')
        if df.empty:
            raise ValueError('工作表为空')

        rows_data = self._df_to_rows(df)
        max_cols = len(df.columns)
        merged_map = {}
        if merged_ranges:
            # ``merged_ranges`` carry original 1-based workbook row numbers,
            # but ``_df_to_rows`` re-indexes the surviving rows sequentially
            # (header -> rendered row 1, data rows -> 2, 3, ...). Any fully
            # empty row removed by ``dropna`` shifts every later row up, so
            # the workbook coordinates no longer line up with the rendered
            # table. Remap them before building ``merged_map`` or merges
            # after a blank row would render at the wrong row (or swallow
            # the wrong cell's value).
            merged_ranges = self._remap_merged_rows(merged_ranges, df)
            merged_map = self._build_merged_map(merged_ranges)
            mc = max((key[1] for key in merged_map), default=0)
            max_cols = max(max_cols, mc)
            # Bug guard: a merge whose legs extend past the available
            # rows would render a rowspan that no tbody row covers,
            # so the browser silently clamps the cell and the visual
            # shape becomes wrong. ``pd.read_excel`` silently drops
            # trailing fully-empty workbook rows, so the dropna-based
            # filter above cannot detect this case. Drop the merge
            # entirely when any leg is past the rendered rows.
            rows_after_dropna = len(rows_data)
            for key, info in list(merged_map.items()):
                if info.max_row > rows_after_dropna:
                    master_key = (info.min_row, info.min_col)
                    for leg in self._iter_merge_legs(info):
                        merged_map.pop(leg, None)
                    merged_map.pop(master_key, None)


        stem = stem_override or clean_filename(Path(input_path).stem)
        sn_clean = sn_override or clean_filename(sheet_name)
        source_name = Path(input_path).name

        if output_fmt == 'html':
            return self._write_html(rows_data, merged_map, max_cols, sheet_name, stem, sn_clean, source_name, parent)
        elif output_fmt == 'md':
            return self._write_md(df, rows_data, merged_map, max_cols, sheet_name, stem, sn_clean, source_name, parent, enhanced_md)
        elif output_fmt == 'json':
            return self._write_json(rows_data, merged_map, sheet_name, stem, sn_clean, parent)
        raise ValueError(f'不支持的输出格式： {output_fmt}')

    # ── HTML output ──────────────────────────────────────────────

    def _write_html(
        self, rows_data: list, merged_map: dict, max_cols: int,
        sheet_name: str, stem: str, sn_clean: str, source_name: str, parent: Path,
    ) -> tuple[str, str]:
        html_table = self._build_html_table(rows_data, merged_map, max_cols, sheet_name, source_name)
        full_html = _html_document(html_table, sheet_name)
        output_name = f'{stem}_{sn_clean}.html'
        output_path = str(parent / output_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self._export(full_html, 'html'))
        return output_name, output_path

    def _build_html_table(
        self, rows_data: list, merged_map: dict, max_cols: int, sheet_name: str,
        source_name: str,
    ) -> str:
        header_row = rows_data[0] if rows_data else []
        html_rows = []
        header_cells = []
        # Cells in the body covered by a header-anchored vertical merge.
        # A ``<th rowspan>`` cannot reach from <thead> into <tbody>, so such
        # merges are flattened: the header renders a normal cell and the
        # covered body cells become placeholders, keeping the table aligned.
        header_covered: dict = {}

        for col_pos in range(1, max_cols + 1):
            h = header_row[col_pos - 1] if col_pos <= len(header_row) else ''
            key = (1, col_pos)
            cell_info = merged_map.get(key)
            if cell_info and cell_info.is_master:
                if cell_info.max_row > 1:
                    for r in range(1, cell_info.rowspan):
                        for c in range(cell_info.colspan):
                            header_covered[(1 + r, col_pos + c)] = True
                    # A header-anchored merge that also spans columns still
                    # needs its colspan so the thead keeps the same column
                    # count as the tbody; rowspan alone cannot cross the
                    # thead/tbody boundary, so it is flattened (see above).
                    span = (
                        f' colspan="{cell_info.colspan}"'
                        f' data-colspan="{cell_info.colspan}"'
                        if cell_info.colspan > 1 else ''
                    )
                    cell_value = html_mod.escape(str(h)) if h is not None and str(h).strip() != '' else '&nbsp;'
                    header_cells.append(f'<th scope="col" data-row="1" data-col="{col_pos}"{span}>{cell_value}</th>')
                else:
                    attrs = self._cell_attrs(1, col_pos, cell_info)
                    cell_value = html_mod.escape(str(h)) if h is not None and str(h).strip() != '' else '&nbsp;'
                    header_cells.append(f'<th{attrs}>{cell_value}</th>')
            elif cell_info and not cell_info.is_master:
                continue
            else:
                cell_value = html_mod.escape(str(h)) if h is not None and str(h).strip() != '' else '&nbsp;'
                header_cells.append(f'<th scope="col" data-row="1" data-col="{col_pos}">{cell_value}</th>')
        html_rows.append('<tr data-row="1">' + ''.join(header_cells) + '</tr>')

        row_spans: dict = {}
        for row_idx, row_data in enumerate(rows_data[1:], start=2):
            cells = []
            for col in range(1, max_cols + 1):
                if header_covered.get((row_idx, col)):
                    cells.append('<td>&nbsp;</td>')
                    continue
                if row_spans.get((row_idx, col), False):
                    continue
                value = row_data[col - 1] if col <= len(row_data) else ''
                key = (row_idx, col)
                cell_info = merged_map.get(key)
                if cell_info and not cell_info.is_master:
                    continue
                if cell_info and cell_info.is_master:
                    attrs = self._cell_attrs(row_idx, col, cell_info)
                    for r in range(1, cell_info.rowspan):
                        for c in range(cell_info.colspan):
                            row_spans[(row_idx + r, col + c)] = True
                    cell_value = html_mod.escape(str(value)) if value is not None and str(value).strip() != '' else '&nbsp;'
                    cells.append(f'<td{attrs}>{cell_value}</td>')
                else:
                    cell_value = html_mod.escape(str(value)) if value is not None and str(value).strip() != '' else '&nbsp;'
                    cells.append(f'<td>{cell_value}</td>')
            html_rows.append(f'<tr data-row="{row_idx}">' + ''.join(cells) + '</tr>')

        source_name = html_mod.escape(source_name)
        return (
            f'<table border="1" class="excel-data" id="data-table" '
            f'data-sheet="{html_mod.escape(sheet_name)}"'
            f' data-source="{source_name}">\n<thead>\n{html_rows[0]}\n</thead>\n<tbody>\n'
            + '\n'.join(html_rows[1:]) + '\n</tbody>\n</table>'
        )

    @staticmethod
    def _cell_attrs(row_idx: int, col_idx: int, info: MergeInfo) -> str:
        parts = [f'data-row="{row_idx}"', f'data-col="{col_idx}"']
        if info.rowspan > 1:
            parts.append(f'rowspan="{info.rowspan}"')
            parts.append(f'data-rowspan="{info.rowspan}"')
        if info.colspan > 1:
            parts.append(f'colspan="{info.colspan}"')
            parts.append(f'data-colspan="{info.colspan}"')
        return ' ' + ' '.join(parts)

    # ── Markdown output (via markitdown) ─────────────────────────

    def _write_md(
        self, df: pd.DataFrame, rows_data: list, merged_map: dict, max_cols: int,
        sheet_name: str, stem: str, sn_clean: str, source_name: str, parent: Path, enhanced: bool,
    ) -> tuple[str, str]:
        if enhanced:
            md_content = self._generate_md_via_html(rows_data, merged_map, max_cols, sheet_name)
        else:
            md_content = self._generate_md_standard(df, sheet_name)

        md_content = (
            f'<!-- source: {source_name} | sheet: {sheet_name}'
            f' | rows: {len(rows_data) - 1} | cols: {max_cols} -->\n\n{md_content}'
        )
        output_name = f'{stem}_{sn_clean}.md'
        output_path = str(parent / output_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self._export(md_content, 'md'))
        return output_name, output_path

    def _generate_md_standard(self, df: pd.DataFrame, sheet_name: str) -> str:
        html_content = df.to_html(index=False, na_rep='')
        return html_to_md(self._escape_table_cells(html_content))

    @staticmethod
    def _escape_table_cells(html_content: str) -> str:
        """Apply escape_md_cell to every th / td text node.

        df.to_html escapes newlines as a literal backslash-n
        sequence (a 2-char run) so HTML renderers display them as
        text. For Markdown output that is wrong: we want the cell
        to render a real line break (``<br>``). Undo pandas'
        escaping before applying escape_md_cell, so the Markdown
        pipeline sees an actual newline and emits ``<br>``.

        Only the standard MD path runs through here. The enhanced
        MD path feeds rows_data directly into the HTML builder,
        where newlines are never escaped by pandas.
        """
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        for cell in soup.find_all(['th', 'td']):
            current = cell.get_text()
            unescaped = current.replace(chr(92) + chr(110), chr(10))
            escaped = escape_md_cell(unescaped)
            if escaped != current:
                cell.string = escaped
        return str(soup)

    def _generate_md_via_html(
        self, rows_data: list, merged_map: dict, max_cols: int, sheet_name: str,
    ) -> str:
        row_spans: dict = {}
        html_rows = ['<table>']
        for row_idx, row_data in enumerate(rows_data, start=1):
            cells = []
            for col in range(1, max_cols + 1):
                span_val = row_spans.get((row_idx, col))
                if span_val is not None:
                    cells.append(span_val)
                    continue
                value = row_data[col - 1] if col <= len(row_data) else ''
                key = (row_idx, col)
                cell_info = merged_map.get(key)
                tag = 'th' if row_idx == 1 else 'td'
                if cell_info and not cell_info.is_master:
                    cells.append(f'<{tag}>&nbsp;</{tag}>')
                    continue
                if cell_info and cell_info.is_master:
                    if value is not None and str(value).strip() != '':
                        # escape_md_cell turns a real newline into a literal
                        # ``<br>``; escaping afterwards keeps it as text
                        # (``&lt;br&gt;``) so the intermediate HTML parsed by
                        # ``html_to_md`` does not turn it into a real element
                        # that markdownify then collapses to a space. This
                        # mirrors the standard-MD path in ``_escape_table_cells``.
                        cell_value = html_mod.escape(escape_md_cell(str(value)))
                    else:
                        cell_value = '&nbsp;'
                    cells.append(f'<{tag}>{cell_value}</{tag}>')
                    for r in range(cell_info.rowspan):
                        for c in range(cell_info.colspan):
                            if r == 0 and c == 0:
                                continue
                            row_spans[(row_idx + r, col + c)] = f'<{tag}>{cell_value}</{tag}>'
                else:
                    if value is not None and str(value).strip() != '':
                        cell_value = html_mod.escape(escape_md_cell(str(value)))
                    else:
                        cell_value = '&nbsp;'
                    cells.append(f'<{tag}>{cell_value}</{tag}>')
            html_rows.append('<tr>' + ''.join(cells) + '</tr>')
        html_rows.append('</table>')
        return html_to_md('\n'.join(html_rows))

    # ── JSON output ──────────────────────────────────────────────

    def _write_json(
        self, rows_data: list, merged_map: dict, sheet_name: str,
        stem: str, sn_clean: str, parent: Path,
    ) -> tuple[str, str]:
        data = _generate_json_data(rows_data, merged_map, sheet_name)
        output_name = f'{stem}_{sn_clean}.json'
        output_path = str(parent / output_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self._export(data, 'json'))
        return output_name, output_path


# ── Module-level helpers ─────────────────────────────────────────


def _html_document(table_html: str, sheet_name: str) -> str:
    return (
        f'<!DOCTYPE html>\n<html lang="zh-CN" data-exported-by="DocConvert" '
        f'data-sheet-name="{html_mod.escape(sheet_name)}">\n'
        f'<head>\n    <meta charset="UTF-8">\n    '
        f'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        f'    <title>{html_mod.escape(sheet_name)}</title>\n'
        f'    <style>\n'
        f'        body {{ font-family: Arial, sans-serif; margin: 20px; }}\n'
        f'        table {{ border-collapse: collapse; width: 100%; }}\n'
        f'        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; white-space: pre-wrap; }}\n'
        f'        th {{ background-color: #4CAF50; color: white; }}\n'
        f'        tr:nth-child(even) {{ background-color: #f2f2f2; }}\n'
        f'        tbody tr {{ cursor: pointer; }}\n'
        f'        tbody tr:hover {{ background-color: #e0e0e0; }}\n'
        f'    </style>\n</head>\n<body>\n{table_html}\n</body>\n</html>'
    )


def _generate_json_data(rows_data: list, merged_map: dict, sheet_name: str) -> str:
    if not rows_data or len(rows_data) < 2:
        return json.dumps(
            {'metadata': {'sheet': sheet_name, 'total_rows': 0, 'total_columns': 0, 'headers': [], 'merged_cells': []}, 'data': []},
            ensure_ascii=False, indent=2,
        )

    headers = list(rows_data[0])
    for c in range(len(headers)):
        if not headers[c]:
            for r in range(1, len(rows_data)):
                if c < len(rows_data[r]) and rows_data[r][c]:
                    headers[c] = rows_data[r][c]
                    break

    merged_cells_info = [
        {'row': k[0], 'col': k[1], 'rowspan': v.rowspan, 'colspan': v.colspan}
        for k, v in merged_map.items() if v.is_master
    ]

    records = []
    for row_idx, row_data in enumerate(rows_data[1:], start=2):
        record: dict = {'_row': row_idx, '_cells': {}}
        for col_idx, value in enumerate(row_data, start=1):
            key = (row_idx, col_idx)
            cell_info = merged_map.get(key)
            header_name = headers[col_idx - 1] if col_idx <= len(headers) else f'col_{col_idx}'

            if cell_info and not cell_info.is_master:
                mr, mc = cell_info.min_row, cell_info.min_col
                value = rows_data[mr - 1][mc - 1] if mr <= len(rows_data) and mc <= len(rows_data[mr - 1]) else ''

            cell_data: dict = {'value': value, 'col': col_idx, 'header': header_name}
            if cell_info and cell_info.is_master:
                cell_data.update(rowspan=cell_info.rowspan, colspan=cell_info.colspan, merged=True)
            elif cell_info and not cell_info.is_master:
                cell_data.update(merged=True, skipped=True)
            else:
                cell_data['merged'] = False
            record['_cells'][str(col_idx)] = cell_data
            record[f"{header_name}_col{col_idx}"] = value
        records.append(record)

    return json.dumps(
        {
            'metadata': {
                'sheet': sheet_name, 'total_rows': len(records),
                'total_columns': len(headers), 'headers': headers,
                'merged_cells': merged_cells_info,
            },
            'data': records,
        },
        ensure_ascii=False, indent=2,
    )
```

---

## `docconvert/converters/word.py`

```python
from __future__ import annotations

import html as html_mod
import io
from pathlib import Path
from typing import Optional

import mammoth
from docx import Document as DocxDocument

from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter
from docconvert.cleaners import WordMdCleaner
from docconvert.utils import clean_filename, html_to_md


class WordConverter(BaseConverter):

    def __init__(self, config: Optional[AppConfig] = None):
        super().__init__(config)
        self.md_cleaner = WordMdCleaner(self.config)

    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        enhanced_md = kwargs.get('enhanced_md', False)
        stem_override = kwargs.get('stem_override')
        if kwargs.get('sheets') is not None:
            raise TypeError('WordConverter does not accept "sheets" parameter')
        results: list[tuple[str, str]] = []
        errors: list[tuple[str, str]] = []
        if self.cancel_check and self.cancel_check():
            return results, errors
        try:
            name, path = self._convert_word(
                input_path, output_fmt, parent, enhanced_md, stem_override
            )
            results.append((name, path))
        except Exception as e:
            self.logger.error("Word转换失败： %s", str(e))
            errors.append((Path(input_path).name, str(e)))
        return results, errors

    def _convert_word(
        self, input_path: str, output_fmt: str, parent: Path,
        enhanced_md: bool = False, stem_override: Optional[str] = None,
    ) -> tuple[str, str]:
        self.logger.info("读取文档： %s", input_path)
        doc = DocxDocument(input_path)
        stem = stem_override or clean_filename(Path(input_path).stem)

        self._strip_headers_footers(doc)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        content = ""
        paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])

        if output_fmt == 'html':
            content = self._to_html(buf, input_path)
        elif output_fmt == 'md':
            content = self._to_md(buf, input_path, paragraph_count, enhanced_md)
        elif output_fmt == 'json':
            content = self._to_json(buf, input_path, paragraph_count)
        else:
            raise ValueError(f'不支持的输出格式： {output_fmt}')

        output_name = f'{stem}_{clean_filename("doc")}.{output_fmt}'
        output_path_full = str(parent / output_name)
        with open(output_path_full, 'w', encoding='utf-8') as f:
            f.write(self._export(content, output_fmt))
        return output_name, output_path_full

    @staticmethod
    def _strip_headers_footers(doc: DocxDocument):
        body_xml = doc.element.body
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        for sect_pr in body_xml.findall(f'.//{ns}sectPr'):
            for child in list(sect_pr):
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag in ('headerReference', 'footerReference'):
                    sect_pr.remove(child)

    def _to_html(self, buf: io.BytesIO, input_path: str) -> str:
        result = mammoth.convert_to_html(buf)
        content = result.value
        title = Path(input_path).stem
        return (
            f'<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n'
            f'    <meta charset="UTF-8">\n'
            f'    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
            f'    <title>{html_mod.escape(title)}</title>\n'
            f'    <style>\n'
            f'        body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; line-height: 1.6; }}\n'
            f'        table {{ border-collapse: collapse; width: 100%; }}\n'
            f'        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; white-space: pre-wrap; }}\n'
            f'        img {{ max-width: 100%; }}\n'
            f'    </style>\n</head>\n<body>\n{content}\n</body>\n</html>'
        )

    def _to_md(self, buf: io.BytesIO, input_path: str, paragraph_count: int, enhanced_md: bool = False) -> str:
        if enhanced_md:
            result = mammoth.convert_to_html(buf)
            content = html_to_md(result.value)
        else:
            result = mammoth.convert_to_markdown(buf)
            content = result.value
        content = self.md_cleaner.clean(content)
        header = (
            f'<!-- source: {Path(input_path).name}'
            f' | paragraphs: {paragraph_count} -->\n\n'
        )
        return header + content

    def _to_json(self, buf: io.BytesIO, input_path: str, paragraph_count: int) -> dict:
        result = mammoth.convert_to_html(buf)
        return {
            'metadata': {
                'source': Path(input_path).name,
                'format': 'docx',
                'paragraphs': paragraph_count,
            },
            'content': result.value,
        }
```

---

## `docconvert/exporters/__init__.py`

```python
from docconvert.exporters.base import BaseExporter
from docconvert.exporters.html import HtmlExporter
from docconvert.exporters.markdown import MarkdownExporter
from docconvert.exporters.json_exporter import JsonExporter

__all__ = [
    "BaseExporter",
    "HtmlExporter",
    "MarkdownExporter",
    "JsonExporter",
    "get_exporter",
]


def get_exporter(fmt: str) -> BaseExporter:
    if fmt == 'html':
        return HtmlExporter()
    elif fmt == 'md':
        return MarkdownExporter()
    elif fmt == 'json':
        return JsonExporter()
    raise ValueError(f'不支持的输出格式： {fmt}')
```

---

## `docconvert/exporters/base.py`

```python
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Optional

from docconvert.config import AppConfig, DEFAULT_CONFIG


class BaseExporter(ABC):

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG

    @abstractmethod
    def export(self, data: Any, **kwargs) -> str:
        ...
```

---

## `docconvert/exporters/html.py`

```python
from __future__ import annotations

from typing import Any

from docconvert.exporters.base import BaseExporter


class HtmlExporter(BaseExporter):

    def export(self, data: Any, **kwargs) -> str:
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            return data.get('content', str(data))
        return str(data)
```

---

## `docconvert/exporters/json_exporter.py`

```python
from __future__ import annotations

from typing import Any

from docconvert.exporters.base import BaseExporter


class JsonExporter(BaseExporter):

    def export(self, data: Any, **kwargs) -> str:
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            import json
            return json.dumps(data, ensure_ascii=False, indent=2)
        return str(data)
```

---

## `docconvert/exporters/markdown.py`

```python
from __future__ import annotations

from typing import Any

from docconvert.exporters.base import BaseExporter


class MarkdownExporter(BaseExporter):

    def export(self, data: Any, **kwargs) -> str:
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            return data.get('content', str(data))
        return str(data)
```

---

## `docconvert/gui/__init__.py`

```python
from docconvert.gui.app import DocConvertApp

__all__ = [
    "DocConvertApp",
]
```

---

## `docconvert/gui/app.py`

```python
from __future__ import annotations

import queue
import tkinter as tk
from pathlib import Path
from threading import Thread
from tkinter import ttk, filedialog, messagebox
from typing import Optional

from docx import Document as DocxDocument

from docconvert.config import DEFAULT_CONFIG, AppConfig
from docconvert.controller import ConversionController
from docconvert.logger import get_logger
from docconvert.models import ProgressEvent
from docconvert.utils import decode_text, get_excel_sheet_names


ALL_EXTS = {'.xlsx', '.xls', '.docx', '.doc'}
ALL_EXTS_LIST = sorted(ALL_EXTS)
ALL_EXTS_PATTERN = ' '.join(f'*{e}' for e in ALL_EXTS_LIST)


# ── Modern color palette ──────────────────────────────────────────────
COLORS = {
    'bg':           '#f0f2f5',
    'card_bg':      '#ffffff',
    'accent':       '#2962ff',
    'accent_hover': '#1e4bd8',
    'accent_light': '#e3edff',
    'success':      '#2e7d32',
    'error':        '#c62828',
    'text':         '#1a1a2e',
    'text_sec':     '#5f6368',
    'border':       '#dadce0',
    'preview_bg':   '#fafbfc',
    'listbox_bg':   '#ffffff',
    'listbox_sel':  '#e3edff',
    'title_bar':    '#2962ff',
}


class DocConvertApp:

    def __init__(self, root):
        self.root = root
        self.logger = get_logger()
        self.controller = ConversionController(DEFAULT_CONFIG)

        self.root.title('DocConvert - 文档转换工具')
        self.root.geometry('780x760')
        self.root.resizable(True, True)
        self.root.minsize(680, 600)
        self.root.configure(bg=COLORS['bg'])

        # Set window icon
        self._set_window_icon()

        self.input_file = tk.StringVar()
        self.file_type: Optional[str] = None
        self.file_paths: list[str] = []
        self.sheet_names: list[str] = []
        self.selected_sheet = tk.StringVar()
        self.output_format = tk.StringVar(value='html')
        self.enhanced_md = tk.BooleanVar(value=False)
        self.clean_page_numbers = tk.BooleanVar(value=True)
        self.clean_dup_headers = tk.BooleanVar(value=True)
        self.clean_empty_lines = tk.BooleanVar(value=True)
        self.clean_normalize_spaces = tk.BooleanVar(value=True)
        self.output_dir = tk.StringVar()
        self.convert_all = tk.BooleanVar(value=False)
        self._call_queue: queue.Queue = queue.Queue()
        self._destroying = False
        self._create_styles()
        self._create_widgets()
        self._create_menu()
        self._setup_hover_effects()
        self.root.update_idletasks()
        self.root.update_idletasks()
        req_h = self.root.winfo_reqheight()
        if req_h > 760:
            self.root.geometry(f'780x{req_h + 20}')
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        self.root.after(50, self._poll_pending)

    def _set_window_icon(self):
        """Set the window icon from the .ico file next to the exe or source."""
        icon_candidates = [
            Path(__file__).parent.parent.parent / 'dist' / 'DocConvert.ico',
            Path(__file__).parent.parent.parent / 'DocConvert.ico',
        ]
        for icon_path in icon_candidates:
            if icon_path.exists():
                try:
                    self.root.iconbitmap(str(icon_path))
                    return
                except tk.TclError:
                    pass
        # Fallback: try to generate a photo icon programmatically
        try:
            self._set_photo_icon()
        except Exception:
            pass

    def _set_photo_icon(self):
        """Create a small blue icon as a PhotoImage fallback."""
        size = 32
        img = tk.PhotoImage(width=size, height=size)
        # Draw a simple blue rounded-square icon
        for y in range(size):
            for x in range(size):
                # Check if inside rounded rectangle
                margin = 2
                r = 4
                ix, iy = x - margin, y - margin
                sx, sy = size - 2 * margin, size - 2 * margin
                in_rect = (margin <= x < size - margin) and (margin <= y < size - margin)
                if in_rect:
                    # Corner rounding check
                    corners = [
                        (r, r), (sx - r, r), (r, sy - r), (sx - r, sy - r)
                    ]
                    in_corner = False
                    for cx, cy in corners:
                        dx, dy = ix - cx, iy - cy
                        if dx < 0 and dy < 0 and dx * dx + dy * dy > r * r:
                            in_corner = True
                            break
                    if not in_corner:
                        img.put('#2962ff', (x, y))
        self._icon_photo = img
        self.root.tk.call('wm', 'iconphoto', self.root._w, img)

    def _create_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # ── Frame styles ──
        style.configure('Card.TFrame', background=COLORS['card_bg'], relief='flat')
        style.configure('Main.TFrame', background=COLORS['bg'])
        style.configure('TitleBar.TFrame', background=COLORS['title_bar'])
        style.configure('Bottom.TFrame', background=COLORS['bg'])
        style.configure('Input.TFrame', background=COLORS['card_bg'])

        # ── Label styles ──
        style.configure('AppTitle.TLabel',
                        font=('Microsoft YaHei UI', 20, 'bold'),
                        background=COLORS['title_bar'],
                        foreground='#ffffff')
        style.configure('AppSubtitle.TLabel',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['title_bar'],
                        foreground='#b3c6ff')
        style.configure('CardTitle.TLabel',
                        font=('Microsoft YaHei UI', 10, 'bold'),
                        background=COLORS['card_bg'],
                        foreground=COLORS['text'])
        style.configure('Field.TLabel',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['card_bg'],
                        foreground=COLORS['text_sec'])
        style.configure('Status.TLabel',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['bg'],
                        foreground=COLORS['text_sec'])
        style.configure('StatusSuccess.TLabel',
                        font=('Microsoft YaHei UI', 9, 'bold'),
                        background=COLORS['bg'],
                        foreground=COLORS['success'])
        style.configure('StatusError.TLabel',
                        font=('Microsoft YaHei UI', 9, 'bold'),
                        background=COLORS['bg'],
                        foreground=COLORS['error'])

        # ── Button styles ──
        style.configure('Accent.TButton',
                        font=('Microsoft YaHei UI', 10, 'bold'),
                        background=COLORS['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        padding=(20, 8))
        style.map('Accent.TButton',
                  background=[('active', COLORS['accent_hover']),
                              ('disabled', '#b0bec5')])

        style.configure('Secondary.TButton',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['card_bg'],
                        foreground=COLORS['text'],
                        borderwidth=1,
                        relief='solid',
                        padding=(12, 5))
        style.map('Secondary.TButton',
                  background=[('active', COLORS['accent_light'])])

        style.configure('Small.TButton',
                        font=('Microsoft YaHei UI', 8),
                        padding=(8, 3))

        # ── Entry styles ──
        style.configure('Modern.TEntry',
                        font=('Microsoft YaHei UI', 9),
                        borderwidth=1,
                        relief='solid',
                        padding=5)

        # ── LabelFrame styles ──
        style.configure('Card.TLabelframe',
                        background=COLORS['card_bg'],
                        foreground=COLORS['text'],
                        borderwidth=1,
                        relief='solid',
                        font=('Microsoft YaHei UI', 10, 'bold'))
        style.configure('Card.TLabelframe.Label',
                        font=('Microsoft YaHei UI', 10, 'bold'),
                        background=COLORS['card_bg'],
                        foreground=COLORS['accent'])

        # ── Checkbutton / Radiobutton ──
        style.configure('Modern.TCheckbutton',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['card_bg'],
                        foreground=COLORS['text'])
        style.configure('Modern.TRadiobutton',
                        font=('Microsoft YaHei UI', 9),
                        background=COLORS['card_bg'],
                        foreground=COLORS['text'])

        # ── Progressbar ──
        style.configure('Accent.Horizontal.TProgressbar',
                        troughcolor=COLORS['border'],
                        background=COLORS['accent'],
                        thickness=6)

        # ── Combobox ──
        style.configure('Modern.TCombobox',
                        font=('Microsoft YaHei UI', 9),
                        padding=4)

        # ── Separator ──
        style.configure('Grey.TSeparator', background=COLORS['border'])

    def _setup_hover_effects(self):
        """Add hover color changes to accent buttons."""
        self.convert_btn.bind('<Enter>', lambda e: self.convert_btn.configure(
            style='AccentHover.TButton' if self.convert_btn.cget('state') == 'normal' else 'Accent.TButton'))
        self.convert_btn.bind('<Leave>', lambda e: self.convert_btn.configure(style='Accent.TButton'))

        # Define hover style
        style = ttk.Style()
        style.configure('AccentHover.TButton',
                        font=('Microsoft YaHei UI', 10, 'bold'),
                        background=COLORS['accent_hover'],
                        foreground='#ffffff',
                        borderwidth=0,
                        padding=(20, 8))

    @property
    def is_converting(self) -> bool:
        return self.controller.is_running

    def _on_close(self):
        self._destroying = True
        if self.controller.is_running:
            self.controller.cancel()
        self.root.after(200, self._real_close)

    def _real_close(self):
        if not self.controller.is_running:
            self.root.destroy()
        else:
            self.root.after(100, self._real_close)

    def _create_menu(self):
        menubar = tk.Menu(self.root, bg=COLORS['card_bg'], fg=COLORS['text'],
                          activebackground=COLORS['accent_light'], activeforeground=COLORS['accent'])
        self.root.config(menu=menubar)
        help_menu = tk.Menu(menubar, tearoff=0, bg=COLORS['card_bg'], fg=COLORS['text'],
                            activebackground=COLORS['accent_light'], activeforeground=COLORS['accent'])
        menubar.add_cascade(label='帮助', menu=help_menu)
        help_menu.add_command(label='使用说明', command=self._show_help)
        help_menu.add_separator()
        help_menu.add_command(label='关于', command=self._show_about)

    def _poll_pending(self):
        """Drain the worker -> main callback queue on the Tk main thread.

        Tkinter is not thread-safe: ``root.after`` must only ever be called
        from the main thread. Background threads therefore enqueue callbacks
        here and this poller (scheduled via ``after`` on the main thread)
        runs them.
        """
        if self._destroying:
            return
        while True:
            try:
                fn = self._call_queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn()
            except Exception as e:
                self.logger.warning("UI 回调异常： %s", e)
        self.root.after(50, self._poll_pending)

    def _run_in_thread(self, work, on_done=None):
        def _wrapper():
            try:
                result = work()
            except Exception as e:
                self.logger.warning("Background worker error: %s", e)
                result = None
            if on_done is not None:
                def _dispatch(r=result):
                    on_done(r)
                self._call_queue.put(_dispatch)
        Thread(target=_wrapper, daemon=True).start()

    def _show_help(self):
        help_text = (
            '文档转换工具 使用说明：\n\n'
            '1. 点击"浏览"选择单个文件，或"添加文件"批量选择多个\n'
            '2. 在文件列表中点击可预览各文件内容\n'
            '3. 选择输出格式（HTML/Markdown/JSON）\n'
            '4. 选择输出目录（可选，默认为第一个文件的目录）\n'
            '5. 点击"开始转换"处理列表中的所有文件\n\n'
            '提示：\n'
            '• Excel 支持带合并单元格的表格\n'
            '• Word 转换会自动清除页眉/页脚等非结构化内容\n'
            '• 勾选"增强"可获得更好的 Markdown 效果\n'
        )
        messagebox.showinfo('使用说明', help_text)

    def _show_about(self):
        messagebox.showinfo('关于', 'DocConvert v2.0\n\n文档转换工具\n支持 Excel/Word 格式转 HTML/Markdown/JSON')

    def _create_widgets(self):
        # ── Title bar ──
        title_bar = ttk.Frame(self.root, style='TitleBar.TFrame')
        title_bar.pack(fill=tk.X)
        title_inner = ttk.Frame(title_bar, style='TitleBar.TFrame')
        title_inner.pack(padx=20, pady=(14, 12), anchor=tk.W)
        ttk.Label(title_inner, text='DocConvert', style='AppTitle.TLabel').pack(side=tk.LEFT)
        ttk.Label(title_inner, text='   文档转换工具  v2.0', style='AppSubtitle.TLabel').pack(side=tk.LEFT, padx=(8, 0))

        # ── Main content area ──
        main_frame = ttk.Frame(self.root, style='Main.TFrame', padding=(15, 10, 15, 5))
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._build_input_card(main_frame)
        self._build_file_list_card(main_frame)
        self._build_format_card(main_frame)
        self._build_preview_card(main_frame)
        self._build_bottom_bar(main_frame)

    def _build_input_card(self, parent):
        card = ttk.LabelFrame(parent, text=' 文件选择 ', style='Card.TLabelframe', padding=12)
        card.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(card, text='输入文件:', style='Field.TLabel').grid(row=0, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        path_frame = ttk.Frame(card, style='Input.TFrame')
        path_frame.grid(row=0, column=1, sticky=tk.EW, pady=6)
        self.path_entry = ttk.Entry(path_frame, textvariable=self.input_file,
                                    font=('Microsoft YaHei UI', 9), style='Modern.TEntry')
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(path_frame, text='浏览', style='Secondary.TButton',
                   command=self._browse_file, width=8).pack(side=tk.LEFT)

        ttk.Label(card, text='工作表:', style='Field.TLabel').grid(row=1, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        sheet_frame = ttk.Frame(card, style='Input.TFrame')
        sheet_frame.grid(row=1, column=1, sticky=tk.EW, pady=6)
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet,
                                        state='readonly', width=35,
                                        font=('Microsoft YaHei UI', 9), style='Modern.TCombobox')
        self.sheet_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(card, text='输出目录:', style='Field.TLabel').grid(row=2, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        outdir_frame = ttk.Frame(card, style='Input.TFrame')
        outdir_frame.grid(row=2, column=1, sticky=tk.EW, pady=6)
        self.outdir_entry = ttk.Entry(outdir_frame, textvariable=self.output_dir,
                                      font=('Microsoft YaHei UI', 9), style='Modern.TEntry')
        self.outdir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(outdir_frame, text='选择', style='Secondary.TButton',
                   command=self._browse_output_dir, width=8).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(outdir_frame, text='默认', style='Small.TButton',
                   command=self._use_default_dir, width=5).pack(side=tk.LEFT)

        self.batch_check = ttk.Checkbutton(card, text='转换所有工作表',
                                           variable=self.convert_all, command=self._toggle_batch,
                                           style='Modern.TCheckbutton')
        self.batch_check.grid(row=3, column=1, sticky=tk.W, pady=2)

        card.columnconfigure(1, weight=1)

    def _build_file_list_card(self, parent):
        card = ttk.LabelFrame(parent, text=' 文件列表 ', style='Card.TLabelframe', padding=8)
        card.pack(fill=tk.X, pady=(0, 8))

        btn_row = ttk.Frame(card, style='Input.TFrame')
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_row, text='添加文件', style='Secondary.TButton',
                   command=self._add_files, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text='移除选中', style='Secondary.TButton',
                   command=self._remove_file, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text='清空', style='Small.TButton',
                   command=self._clear_files, width=6).pack(side=tk.LEFT, padx=2)

        list_frame = ttk.Frame(card, style='Input.TFrame')
        list_frame.pack(fill=tk.X)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(list_frame, height=4, font=('Consolas', 9),
                                       bg=COLORS['listbox_bg'], fg=COLORS['text'],
                                       selectbackground=COLORS['listbox_sel'],
                                       selectforeground=COLORS['accent'],
                                       borderwidth=1, relief='solid',
                                       highlightthickness=1, highlightcolor=COLORS['accent'],
                                       yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(fill=tk.X, expand=True)
        self.file_listbox.bind('<<ListboxSelect>>', self._on_select_file)

    def _build_format_card(self, parent):
        card = ttk.LabelFrame(parent, text=' 输出格式 ', style='Card.TLabelframe', padding=12)
        card.pack(fill=tk.X, pady=(0, 8))

        radio_row = ttk.Frame(card, style='Input.TFrame')
        radio_row.pack(fill=tk.X, pady=(0, 6))

        # Custom-styled radio buttons
        for text, val in [('HTML', 'html'), ('Markdown', 'md'), ('JSON', 'json')]:
            rb = ttk.Radiobutton(radio_row, text=text, variable=self.output_format,
                                 value=val, command=self._update_enhanced_state,
                                 style='Modern.TRadiobutton')
            rb.pack(side=tk.LEFT, padx=20)

        self.enhanced_check = ttk.Checkbutton(
            card, text='增强 Markdown 输出（更好的格式）',
            variable=self.enhanced_md, state='disabled',
            style='Modern.TCheckbutton'
        )
        self.enhanced_check.pack(anchor=tk.W, pady=(4, 0))

        self.cleaning_frame = ttk.LabelFrame(card, text=' Markdown 清洗 (Word→MD) ',
                                             style='Card.TLabelframe', padding=8)
        self.cleaning_frame.pack(fill=tk.X, pady=(8, 0))

        checks_data = [
            (self.clean_page_numbers, '移除页码 ([1] / 第N页 / Page X 等）', 0, 0),
            (self.clean_dup_headers, '移除重复页眉', 0, 1),
            (self.clean_empty_lines, '移除多余空行', 1, 0),
            (self.clean_normalize_spaces, '合并多余空白', 1, 1),
        ]
        for var, text, r, c in checks_data:
            cb = ttk.Checkbutton(self.cleaning_frame, text=text, variable=var,
                                 state='disabled', style='Modern.TCheckbutton')
            cb.grid(row=r, column=c, sticky=tk.W, padx=6, pady=2)

        self.cleaning_checks = [
            self.clean_page_numbers, self.clean_dup_headers,
            self.clean_empty_lines, self.clean_normalize_spaces,
        ]

    def _build_preview_card(self, parent):
        card = ttk.LabelFrame(parent, text=' 预览 ', style='Card.TLabelframe', padding=8)
        card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.preview_text = tk.Text(
            card, height=6, font=('Consolas', 9),
            bg=COLORS['preview_bg'], fg=COLORS['text'],
            insertbackground=COLORS['text'],
            selectbackground=COLORS['accent_light'],
            borderwidth=0, relief='flat',
            highlightthickness=1, highlightcolor=COLORS['border'],
            state='disabled', padx=8, pady=6
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True)

    def _build_bottom_bar(self, parent):
        bottom = ttk.Frame(parent, style='Bottom.TFrame')
        bottom.pack(fill=tk.X, pady=(0, 5))

        # Left: status
        self.status_label = ttk.Label(bottom, text='就绪', style='Status.TLabel')
        self.status_label.pack(side=tk.LEFT, padx=(5, 0))

        # Right: progress + convert button
        right_frame = ttk.Frame(bottom, style='Bottom.TFrame')
        right_frame.pack(side=tk.RIGHT)

        self.progress = ttk.Progressbar(right_frame, mode='indeterminate', length=180,
                                        style='Accent.Horizontal.TProgressbar')
        self.progress.pack(side=tk.LEFT, padx=(0, 12), pady=8)

        self.convert_btn = ttk.Button(right_frame, text='开始转换', style='Accent.TButton',
                                      command=self._convert)
        self.convert_btn.pack(side=tk.LEFT, pady=4)

    # ── Existing logic (unchanged) ────────────────────────────────────

    def _update_enhanced_state(self):
        is_md = self.output_format.get() == 'md'
        is_word = self.file_type == 'word'
        self.enhanced_check.configure(state='normal' if is_md else 'disabled')
        if not is_md:
            self.enhanced_md.set(False)
        clean_state = 'normal' if (is_md and is_word) else 'disabled'
        for child in self.cleaning_frame.winfo_children():
            try:
                child.configure(state=clean_state)
            except tk.TclError:
                pass
        if not (is_md and is_word):
            for var in self.cleaning_checks:
                var.set(False)

    def _browse_output_dir(self):
        dirname = filedialog.askdirectory(title='选择输出目录')
        if dirname:
            self.output_dir.set(dirname)

    def _use_default_dir(self):
        self.output_dir.set('')

    def _build_config(self) -> AppConfig:
        from dataclasses import replace
        return replace(
            self.controller.config,
            cleaning_rules={
                "remove_page_numbers": self.clean_page_numbers.get(),
                "remove_duplicate_headers": self.clean_dup_headers.get(),
                "remove_empty_lines": self.clean_empty_lines.get(),
                "normalize_spaces": self.clean_normalize_spaces.get(),
            },
        )

    def _update_preview(self, info_type='info', message=''):
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        colors = {
            'info': COLORS['text'],
            'success': COLORS['success'],
            'error': COLORS['error'],
            'header': COLORS['accent'],
        }
        self.preview_text.insert(tk.END, message, info_type)
        self.preview_text.tag_config('info', foreground=colors['info'])
        self.preview_text.tag_config('success', foreground=colors['success'])
        self.preview_text.tag_config('error', foreground=colors['error'])
        self.preview_text.tag_config('header', font=('Consolas', 9, 'bold'), foreground=colors['header'])
        self.preview_text.configure(state='disabled')

    def _show_preview_file(self, filepath):
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        BUF = self.controller.config.preview_chars
        PREVIEW_LINES = self.controller.config.preview_lines
        MAX_READ = 1024 * 1024
        content = None
        read_truncated = False
        for enc in ('utf-8', 'gbk'):
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    content = f.read(MAX_READ + 1)
                if len(content) > MAX_READ:
                    content = content[:MAX_READ]
                    read_truncated = True
                break
            except UnicodeDecodeError:
                continue
            except OSError as e:
                self.preview_text.insert(tk.END, f'无法预览文件： {filepath}\n{e}', 'error')
                self.preview_text.configure(state='disabled')
                return
        if content is None:
            try:
                with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read(MAX_READ + 1)
                if len(content) > MAX_READ:
                    content = content[:MAX_READ]
                    read_truncated = True
            except OSError as e:
                self.preview_text.insert(tk.END, f'无法预览文件： {filepath}\n{e}', 'error')
                self.preview_text.configure(state='disabled')
                return
        try:
            total_chars = len(content)
            if total_chars == 0:
                total_lines = 0
            elif content.endswith('\n'):
                total_lines = content.count('\n')
            else:
                total_lines = content.count('\n') + 1
            display = content[:BUF]
            display_lines = display.split('\n')
            preview_lines = display_lines[:PREVIEW_LINES]
            preview = '\n'.join(preview_lines)
            truncated = read_truncated or total_chars > BUF or len(display_lines) > PREVIEW_LINES
            self.preview_text.insert(tk.END, f'预览： {Path(filepath).name}\n', 'header')
            total_str = f'共 {total_lines} 行， {total_chars} 字符'
            if truncated:
                total_str += ' （截断）'
            self.preview_text.insert(tk.END, total_str + '\n\n', 'info')
            self.preview_text.insert(tk.END, preview)
            if truncated:
                self.preview_text.insert(tk.END, '\n\n... （内容已截断）')
        except (OSError, UnicodeDecodeError):
            self.preview_text.insert(tk.END, f'无法预览文件： {filepath}', 'error')
        self.preview_text.configure(state='disabled')

    def _load_by_ext(self, filepath):
        if not Path(filepath).exists():
            messagebox.showerror('错误', f'文件不存在： {filepath}')
            return False
        ext = Path(filepath).suffix.lower()
        if ext in ('.xlsx', '.xls'):
            self.file_type = 'excel'
            if ext == '.xls':
                try:
                    import xlrd
                except ImportError:
                    messagebox.showerror('错误', '处理 .xls 文件需要安装 xlrd 库\n\n请运行： pip install xlrd')
                    return False
            self.input_file.set(filepath)
            self._load_sheets(filepath)
            self._update_enhanced_state()
            return True
        elif ext == '.docx':
            self.file_type = 'word'
            self.input_file.set(filepath)
            self._load_word(filepath)
            self._update_enhanced_state()
            return True
        elif ext == '.doc':
            try:
                import textract
            except ImportError:
                messagebox.showerror('错误', '处理 .doc 文件需要安装 textract 库\n\n请运行： pip install textract')
                return False
            self.file_type = 'doc'
            self.input_file.set(filepath)
            self.sheet_combo['values'] = []
            self.sheet_combo.set('')
            self.sheet_combo.configure(state='disabled')
            self.sheet_names = []
            self.selected_sheet.set('')
            self._update_preview('info', f'.doc 文件： {Path(filepath).name}\n（正在后台提取文本，请稍候...)')
            self._load_doc_async(filepath)
            self._update_enhanced_state()
            return True
        else:
            messagebox.showerror('错误', '不支持的文件格式')
            return False

    def _browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[
                ('支持的文件', ALL_EXTS_PATTERN),
                ('Excel 文件', '*.xlsx *.xls'),
                ('Word 文件', '*.docx *.doc'),
                ('所有文件', '*.*'),
            ],
            title='选择文件'
        )
        if filename:
            if Path(filename).suffix.lower() in ALL_EXTS:
                if filename not in self.file_paths:
                    self.file_paths.append(filename)
                self._refresh_file_list()
            self._load_by_ext(filename)

    def _add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[
                ('支持的文件', ALL_EXTS_PATTERN),
                ('Excel 文件', '*.xlsx *.xls'),
                ('Word 文件', '*.docx *.doc'),
                ('所有文件', '*.*'),
            ],
            title='添加文件'
        )
        if not files:
            return
        for f in files:
            if Path(f).suffix.lower() in ALL_EXTS and f not in self.file_paths:
                self.file_paths.append(f)
        self._refresh_file_list()
        if self.file_paths:
            self._load_by_ext(self.file_paths[0])

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.file_paths.pop(idx)
        self._refresh_file_list()
        if self.file_paths:
            new_idx = min(idx, len(self.file_paths) - 1)
            self._suppress_select_event = True
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(new_idx)
            self._load_by_ext(self.file_paths[new_idx])
            self.root.after_idle(
                lambda: setattr(self, '_suppress_select_event', False)
            )
        else:
            self.input_file.set('')
            self._update_preview('info', '文件列表为空')

    def _clear_files(self):
        self.file_paths.clear()
        self.file_listbox.delete(0, tk.END)
        self.input_file.set('')
        self._update_preview('info', '文件列表已清空')

    def _refresh_file_list(self):
        self.file_listbox.delete(0, tk.END)
        names = [Path(f).name for f in self.file_paths]
        dupes = {n for n in names if names.count(n) > 1}
        for f in self.file_paths:
            p = Path(f)
            label = f'{p.parent.name}/{p.name}' if p.name in dupes else p.name
            self.file_listbox.insert(tk.END, label)

    def _on_select_file(self, event):
        if getattr(self, '_suppress_select_event', False):
            return
        sel = self.file_listbox.curselection()
        if sel:
            idx = sel[0]
            if idx < len(self.file_paths):
                self._load_by_ext(self.file_paths[idx])

    def _load_sheets(self, filepath):
        self.sheet_combo.configure(state='readonly')
        try:
            ext = Path(filepath).suffix.lower()
            self.sheet_names = get_excel_sheet_names(filepath, ext)
            self.selected_sheet.set('')
            if self.sheet_names:
                self.sheet_combo['values'] = self.sheet_names
                if len(self.sheet_names) == 1:
                    self.selected_sheet.set(self.sheet_names[0])
                else:
                    self.sheet_combo.current(0)
                    self.selected_sheet.set(self.sheet_names[0])
                info = f'已加载： {len(self.sheet_names)} 个工作表\n\n工作表列表:\n'
                for i, name in enumerate(self.sheet_names[:10]):
                    info += f'  {i + 1}. {name}\n'
                if len(self.sheet_names) > 10:
                    info += f'  ... 还有 {len(self.sheet_names) - 10} 个\n'
                self._update_preview('success', info)
        except Exception as e:
            messagebox.showerror('错误', f'加载文件失败:\n{str(e)}')
            self.sheet_names = []
            self.sheet_combo['values'] = []
            self.sheet_combo.set('')
            self._update_preview('error', f'加载失败： {str(e)}')
        finally:
            if self.convert_all.get():
                self.sheet_combo.configure(state='disabled')

    def _load_word(self, filepath):
        try:
            doc = DocxDocument(filepath)
            paras = [p for p in doc.paragraphs if p.text.strip()]
            tables = doc.tables
            sections = doc.sections
            self.sheet_names = []
            self.selected_sheet.set('')
            self.sheet_combo['values'] = []
            self.sheet_combo.set('')
            self.sheet_combo.configure(state='disabled')
            info = f'Word 文档： {Path(filepath).name}\n'
            info += f'段落数： {len(paras)}\n'
            info += f'表格数： {len(tables)}\n'
            info += f'节数： {len(sections)}\n\n'
            if sections:
                hf_count = 0
                for sec in sections:
                    if sec.header and any(p.text.strip() for p in sec.header.paragraphs):
                        hf_count += 1
                    if sec.footer and any(p.text.strip() for p in sec.footer.paragraphs):
                        hf_count += 1
                if hf_count:
                    info += '含页眉/页脚 （将被清洗）\n'
            if paras:
                info += '\n--- 预览前10段 ---\n'
                for p in paras[:10]:
                    info += p.text[:120] + '\n'
            self._update_preview('success', info)
        except Exception as e:
            messagebox.showerror('错误', f'加载文档失败:\n{str(e)}')
            self._update_preview('error', f'加载失败： {str(e)}')

    def _load_doc_async(self, filepath: str):
        def _work():
            try:
                import textract
                raw = textract.process(filepath)
                text = decode_text(raw)
                lines = text.split('\n')
                preview_lines = [line for line in lines if line.strip()][:20]
                return (
                    f'.doc 文件： {Path(filepath).name}\n'
                    f'共约 {len(lines)} 行\n\n--- 预览 ---\n'
                    + '\n'.join(preview_lines)
                )
            except Exception:
                return f'.doc 文件： {Path(filepath).name}\n（后台文本提取失败）'

        def _on_done(info):
            if self.file_type == 'doc' and self.input_file.get() == filepath:
                self._update_preview('info', info)

        self._run_in_thread(_work, _on_done)

    def _toggle_batch(self):
        if self.convert_all.get():
            self.sheet_combo.configure(state='disabled')
        else:
            self.sheet_combo.configure(state='readonly')

    def _on_progress(self, event: ProgressEvent):
        self._call_queue.put(lambda e=event: self._handle_progress(e))

    def _handle_progress(self, event: ProgressEvent):
        if event.error:
            self.status_label.config(text=event.error, style='StatusError.TLabel')
            self._update_preview('error', event.error)
        elif event.message:
            self.status_label.config(text=event.message, style='Status.TLabel')

        if event.progress > 0:
            # First determinate update: stop the indeterminate animation
            # before switching modes so the bar doesn't keep cycling.
            if str(self.progress['mode']) == 'indeterminate':
                self.progress.stop()
            self.progress['mode'] = 'determinate'
            self.progress['value'] = int(event.progress * 100)

    def _convert(self):
        if self.is_converting:
            return

        output_fmt = self.output_format.get()

        files_to_process = list(self.file_paths)
        if not files_to_process:
            p = self.input_file.get().strip()
            if p:
                files_to_process = [p]

        if not files_to_process:
            messagebox.showerror('错误', '请选择文件')
            return

        out_dir = self.output_dir.get().strip()

        sheets = None
        if not self.convert_all.get():
            sel = self.selected_sheet.get()
            if sel:
                sheets = [sel]

        self.controller.set_config(self._build_config())

        existing = self.controller.check_overwrite_paths(
            files=files_to_process,
            output_fmt=output_fmt,
            output_dir=out_dir if out_dir else None,
            sheets=sheets,
        )
        if existing:
            preview = '\n'.join(f'  \u2022 {Path(p).name}' for p in existing[:10])
            if len(existing) > 10:
                preview += f'\n  ... 还有 {len(existing) - 10} 个'
            proceed = messagebox.askyesno(
                '覆盖确认',
                f'以下 {len(existing)} 个输出文件已存在，将被覆盖:\n{preview}\n\n是否继续？',
            )
            if not proceed:
                return

        self.convert_btn.configure(state='disabled')
        self.status_label.config(text='加载文件中...')
        self.progress.configure(mode='indeterminate', value=0)
        # Animate the bar during the initial loading phase, before the
        # first determinate progress event arrives (which stops it).
        self.progress.start(12)
        self.root.update()

        started = self.controller.convert_files_async(
            files=files_to_process,
            output_fmt=output_fmt,
            output_dir=out_dir if out_dir else None,
            enhanced_md=self.enhanced_md.get(),
            sheets=sheets,
            progress_callback=self._on_progress,
        )

        if not started:
            self.convert_btn.configure(state='normal')
            self.progress.stop()
            self.progress.configure(mode='determinate', value=0)
            self.status_label.config(text='任务已在运行中', style='StatusError.TLabel')
            return

        self._wait_for_done()

    def _wait_for_done(self):
        self._run_in_thread(self.controller.wait_done, self._on_conversion_done)

    def _on_conversion_done(self, _completed: bool = True):
        self.convert_btn.configure(state='normal')
        # Stop any residual indeterminate animation (e.g. a batch that
        # finished before emitting a determinate progress event).
        if str(self.progress['mode']) == 'indeterminate':
            self.progress.stop()
        self.progress['mode'] = 'determinate'

        if self.controller.last_error:
            err = self.controller.last_error
            self.status_label.config(text='转换失败', style='StatusError.TLabel')
            self._update_preview('error', f'转换失败： {err}')
            messagebox.showerror('错误', f'转换失败:\n{err}')
            return

        if not self.controller.was_cancelled:
            self.progress['value'] = 100
        self._show_conversion_results()

    def _show_conversion_results(self):
        results = self.controller.last_results
        all_results = [(n, p) for n, p, e in results if e is None]
        all_errors = [(n, e) for n, p, e in results if e is not None]
        was_cancelled = self.controller.was_cancelled

        if not results:
            self.status_label.config(text='已取消', style='Status.TLabel')
            self._update_preview('info', '转换已取消')
            return

        total = len(results)
        is_multi = total > 1

        if is_multi:
            msg_lines = [f'处理完成： {len(all_results)}/{total} 个输出']
            if all_results:
                msg_lines.append('')
                msg_lines.append('成功:')
                for n, p in all_results[:20]:
                    msg_lines.append(f'  \u2022 {n}')
                if len(all_results) > 20:
                    msg_lines.append(f'  ... 还有 {len(all_results) - 20} 个')
            if all_errors:
                msg_lines.append('')
                msg_lines.append(f'失败 ({len(all_errors)}):')
                for fname, err in all_errors[:10]:
                    msg_lines.append(f'  \u2022 {fname}: {err}')
                if len(all_errors) > 10:
                    msg_lines.append(f'  ... 还有 {len(all_errors) - 10} 个')

            if was_cancelled:
                status_text = f'已取消 ({len(all_results)}/{total})'
                status_style = 'Status.TLabel'
            elif all_results and not all_errors:
                status_text = f'转换完成 ({len(all_results)}/{total})'
                status_style = 'StatusSuccess.TLabel'
            elif all_results and all_errors:
                status_text = f'部分成功 ({len(all_results)}/{total})'
                status_style = 'Status.TLabel'
            else:
                status_text = f'转换失败 (0/{total})'
                status_style = 'StatusError.TLabel'

            self.status_label.config(text=status_text, style=status_style)
            if all_results:
                self._show_preview_file(all_results[0][1])
            messagebox.showinfo('转换结果', '\n'.join(msg_lines))
        elif all_results:
            name, path = all_results[0]
            self.status_label.config(text='转换完成', style='StatusSuccess.TLabel')
            self._show_preview_file(path)
            ext_map = {'html': 'HTML', 'md': 'Markdown', 'json': 'JSON'}
            messagebox.showinfo(
                '成功',
                f'{ext_map.get(self.output_format.get(), self.output_format.get())} 文件已生成:\n{path}'
            )
        else:
            err_msg = all_errors[0][1] if all_errors else '未知错误'
            self.status_label.config(text='转换失败', style='StatusError.TLabel')
            self._update_preview('error', f'转换失败： {err_msg}')
            messagebox.showerror('错误', f'转换失败:\n{err_msg}')
```

---

## `docconvert/logger.py`

```python
from __future__ import annotations

import logging
import sys
from typing import Union


LOGGER_NAME = "docconvert"


def setup_logging(level: Union[int, str] = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)

    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(level)
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(name)s - %(message)s",
            datefmt="%H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    else:
        # A later call (e.g. main.py -> main_cli with --verbose) must
        # raise the already-registered handler's level too; setting only
        # the logger's level is not enough, the handler would keep
        # filtering the messages out.
        for handler in logger.handlers:
            handler.setLevel(level)

    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger(LOGGER_NAME)
```

---

## `docconvert/models/__init__.py`

```python
from docconvert.models.models import (
    MergeInfo,
    ProgressEvent,
)

__all__ = [
    "MergeInfo",
    "ProgressEvent",
]
```

---

## `docconvert/models/models.py`

```python
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class MergeInfo:
    rowspan: int = 1
    colspan: int = 1
    is_master: bool = False
    is_merged: bool = False
    min_row: int = 0
    min_col: int = 0
    max_row: int = 0
    max_col: int = 0


@dataclass
class ProgressEvent:
    message: str = ""
    progress: float = 0.0
    done: bool = False
    error: Optional[str] = None
```

---

## `docconvert/parsers/__init__.py`

```python
from docconvert.parsers.semantic import BaseParser

__all__ = [
    "BaseParser",
]
```

---

## `docconvert/parsers/semantic.py`

```python
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class BaseParser(ABC):
    @abstractmethod
    def parse(self, content: Any, **kwargs) -> Any:
        ...
```

---

## `docconvert/utils/__init__.py`

```python
from docconvert.utils.utils import (
    INVALID_NAMES,
    safe_str,
    clean_filename,
    escape_md_cell,
    html_to_md,
    get_excel_sheet_names,
    decode_text,
    unique_cleaned_suffixes,
)

__all__ = [
    "INVALID_NAMES",
    "safe_str",
    "clean_filename",
    "escape_md_cell",
    "html_to_md",
    "get_excel_sheet_names",
    "decode_text",
    "unique_cleaned_suffixes",
]
```

---

## `docconvert/utils/utils.py`

```python
from __future__ import annotations

import re
from typing import Any

INVALID_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def get_excel_sheet_names(filepath: str, ext: str) -> list[str]:
    """Return the list of sheet names in an Excel workbook.

    ``ext`` is the lowercased suffix (``.xls`` or ``.xlsx``); ``.xls`` is
    read via xlrd and ``.xlsx`` via openpyxl.
    """
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(filepath)
        try:
            return wb.sheet_names()
        finally:
            wb.release_resources()
    from openpyxl import load_workbook
    wb = load_workbook(filepath, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def safe_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and str(value) == 'nan':
        return ""
    try:
        import pandas as pd
        if pd.isna(value):
            return ""
    except (ImportError, TypeError, ValueError):
        pass
    s = str(value)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s


def clean_filename(name: str) -> str:
    name = re.sub(r'[\x00-\x1f]', '', str(name))
    name = re.sub(r'[\/\\:*?"<>|]', "_", name)
    if not name.strip():
        name = "untitled"
    stem = name.split('.')[0].upper()
    if stem in INVALID_NAMES:
        name = f"_{name}"
    if len(name) > 180:
        name = name.encode('utf-8')[:180].decode('utf-8', errors='replace')
    return name


def unique_cleaned_suffixes(names: list[str]) -> list[str]:
    """Return a clean_filename-escaped, batch-unique suffix per name.

    Inputs that differ but clean to the same value (e.g. sheet names
    ``Q"1`` and ``Q<1`` both become ``Q_1``) would otherwise produce
    identical output paths and the later item would silently overwrite
    the earlier one. Subsequent collisions get a numeric suffix:
    ``name``, ``name_2``, ``name_3``...
    """
    used: set[str] = set()
    result: list[str] = []
    for name in names:
        base = clean_filename(name)
        candidate = base
        n = 2
        while candidate in used:
            candidate = f'{base}_{n}'
            n += 1
        used.add(candidate)
        result.append(candidate)
    return result


def decode_text(raw: bytes) -> str:
    """Decode byte output from external tools (textract/antiword).

    Their output encoding depends on the document and the system locale
    (UTF-8 on most systems, GBK on zh-CN Windows, Latin-1 on Western).
    Try UTF-8 first, then GB18030 (a superset of GBK/GB2312), then fall
    back to Latin-1, which can never fail.
    """
    for enc in ('utf-8', 'gb18030', 'latin-1'):
        try:
            s = raw.decode(enc)
            if '\ufffd' not in s:
                return s
        except (UnicodeDecodeError, ValueError):
            continue
    return raw.decode('latin-1', errors='replace')


def escape_md_cell(text: str) -> str:
    """Escape Markdown-significant characters inside a table cell.

    Markdown table cells containing a literal ``|`` break the row
    layout because the pipe is also the column separator. Cells
    containing a newline are joined by ``markdownify`` into a
    single space, losing the visible line break. This helper returns
    a Markdown-safe form of the cell text so the rendered table
    keeps its column count and the line break is preserved as
    ``<br>`` (which survives markdownify pass-through).

    Apply this only to text that will be embedded inside a ``<th>``
    / ``<td>``. Do not apply it to the surrounding HTML structure.
    """
    if not text:
        return text
    return text.replace("|", "\\|").replace("\n", "<br>")

def html_to_md(html_content: str) -> str:
    from bs4 import BeautifulSoup
    import markdownify
    soup = BeautifulSoup(html_content, 'html.parser')
    for tag in soup(['script', 'style']):
        tag.extract()
    return markdownify.markdownify(str(soup), heading_style=markdownify.ATX).strip()
```

---


<!-- snapshot: 30 files, 118,588 chars (regenerate with `python gen_doc.py`) -->
