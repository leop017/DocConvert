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
