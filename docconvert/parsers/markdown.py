"""Markdown parser.

Recovers the headings, paragraphs, lists, fenced code blocks and tables of
a Markdown document. It is *not* a full CommonMark implementation -- it
deliberately handles only the surface syntax that the DocConvert pipeline
emits, so that it is cheap, deterministic, and tolerant of malformed input.
"""

from __future__ import annotations

import re
from typing import Any

from docconvert.parsers.base import BaseParser
from docconvert.parsers.models import Document, Element

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
_FENCE_RE = re.compile(r"^(```|~~~)(.*)$")
_LIST_RE = re.compile(r"^\s*(?:[-*+]|\d+\.)\s+")
_TABLE_DIVIDER_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$")
_TABLE_CELL_RE = re.compile(r"\|")


class MarkdownParser(BaseParser):
    """A pragmatic Markdown parser suitable for clean(ish) inputs."""

    def parse(self, content: Any, **kwargs: Any) -> Document:
        text = self._coerce_to_str(content)
        metadata: dict[str, Any] = {}

        if text.startswith("---"):
            text, metadata = self._strip_front_matter(text)

        elements: list[Element] = []
        lines: list[str] = text.splitlines()
        i = 0
        n = len(lines)
        while i < n:
            line = lines[i]
            if not line.strip():
                i += 1
                continue

            m = _FENCE_RE.match(line)
            if m:
                fence_char = m.group(1)[0]
                lang = m.group(2).strip()
                j = i + 1
                body: list[str] = []
                while j < n and lines[j].lstrip().startswith(fence_char) is False:
                    body.append(lines[j])
                    j += 1
                elements.append(Element(
                    element_type="code",
                    text="\n".join(body),
                    metadata={"language": lang},
                ))
                i = j + 1
                continue

            m = _HEADING_RE.match(line)
            if m:
                level = len(m.group(1))
                title = m.group(2).strip()
                elements.append(Element(
                    element_type="heading",
                    text=title,
                    metadata={"level": level},
                ))
                i += 1
                continue

            if _LIST_RE.match(line):
                j = i
                bullet_lines: list[str] = []
                while j < n and (not lines[j].strip() or _LIST_RE.match(lines[j])):
                    bullet_lines.append(lines[j])
                    j += 1
                elements.append(Element(element_type="list", text="\n".join(bullet_lines)))
                i = j
                continue

            if "|" in line and i + 1 < n and _TABLE_DIVIDER_RE.match(lines[i + 1]):
                header_cells = self._split_row(line)
                j = i + 2
                body_rows: list[str] = []
                while j < n and "|" in lines[j] and not _TABLE_DIVIDER_RE.match(lines[j]):
                    body_rows.append(lines[j])
                    j += 1
                elements.append(Element(
                    element_type="table",
                    text="\n".join([line] + body_rows),
                    metadata={"header": header_cells, "rows": body_rows},
                ))
                i = j
                continue

            j = i
            paragraph: list[str] = []
            while j < n and lines[j].strip() and not _HEADING_RE.match(lines[j]) \
                    and not _FENCE_RE.match(lines[j]) and not _LIST_RE.match(lines[j]):
                paragraph.append(lines[j])
                j += 1
            elements.append(Element(element_type="paragraph", text="\n".join(paragraph)))
            i = j

        flat = "\n\n".join(e.text for e in elements)
        return Document(text=flat, elements=elements, metadata=metadata)

    @staticmethod
    def _coerce_to_str(content: Any) -> str:
        if isinstance(content, str):
            return content
        if isinstance(content, bytes):
            return content.decode("utf-8", errors="replace")
        return str(content)

    @staticmethod
    def _strip_front_matter(text: str) -> tuple[str, dict[str, Any]]:
        lines = text.splitlines()
        if not lines or lines[0].strip() != "---":
            return text, {}
        end = None
        for k in range(1, len(lines)):
            if lines[k].strip() == "---":
                end = k
                break
        if end is None:
            return text, {}
        meta: dict[str, Any] = {}
        for raw in lines[1:end]:
            if ":" not in raw:
                continue
            parts = raw.partition(":")
            key: str = parts[0]
            value: str = parts[2]
            meta[key.strip()] = value.strip()
        body = "\n".join(lines[end + 1:])
        return body, meta

    @staticmethod
    def _split_row(row: str) -> list[str]:
        stripped = row.strip()
        if stripped.startswith("|"):
            stripped = stripped[1:]
        if stripped.endswith("|"):
            stripped = stripped[:-1]
        return [c.strip() for c in stripped.split("|")]
