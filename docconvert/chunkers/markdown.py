"""Header-aware Markdown chunker.

Groups consecutive elements under the same heading (and its descendants)
into a single chunk. This is the most useful strategy for DocConvert
output, because the upstream Markdown files are already split by
``## Section`` / ``### Subsection`` markers.
"""

from __future__ import annotations

from typing import Any

from docconvert.chunkers.base import BaseChunker
from docconvert.parsers.models import Chunk, Element


class MarkdownChunker(BaseChunker):
    """Emit one :class:`Chunk` per heading-subtree."""

    def chunk(self, content: Any, **kwargs: Any) -> list[Chunk]:
        max_chunk_size = int(kwargs.get("max_chunk_size", 0)) or None
        doc = self._coerce_to_document(content)
        if not doc.elements:
            return []

        groups: list[tuple[str, list[Element]]] = []
        current_header = ""
        current_body: list[Element] = []

        for elem in doc.elements:
            if elem.element_type == "heading":
                if current_body or current_header:
                    groups.append((current_header, current_body))
                current_header = elem.text
                current_body = []
            else:
                current_body.append(elem)
        if current_body or current_header:
            groups.append((current_header, current_body))

        out: list[Chunk] = []
        total = len(groups)
        cursor = 0
        for i, (header, body) in enumerate(groups):
            text = self._render(header, body)
            start = doc.text.find(text[:40], cursor) if text else cursor
            if start < 0:
                start = cursor
            end = start + len(text)
            cursor = end
            if max_chunk_size and len(text) > max_chunk_size:
                pieces = self._split_long(text, max_chunk_size)
                for j, piece in enumerate(pieces):
                    out.append(self._make_chunk(
                        text=piece,
                        metadata={**doc.metadata, "header": header, "part": j},
                        index=len(out),
                        total=total * len(pieces),
                        start=start + sum(len(p) for p in pieces[:j]),
                        end=start + sum(len(p) for p in pieces[:j + 1]),
                    ))
            else:
                out.append(self._make_chunk(
                    text=text,
                    metadata={**doc.metadata, "header": header},
                    index=i,
                    total=total,
                    start=start,
                    end=end,
                ))
        return out

    @staticmethod
    def _render(header: str, body: list[Element]) -> str:
        parts = []
        if header:
            parts.append(f"# {header}")
        for elem in body:
            if elem.text:
                parts.append(elem.text)
        return "\n\n".join(parts)

    @staticmethod
    def _split_long(text: str, size: int) -> list[str]:
        if size <= 0:
            return [text]
        out: list[str] = []
        for i in range(0, len(text), size):
            out.append(text[i:i + size])
        return out
