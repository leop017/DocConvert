"""Sentence-boundary chunker.

Greedily accumulates sentences until the running window reaches
``chunk_size`` characters, then emits a chunk and starts a new window that
reuses up to ``chunk_overlap`` characters of context. Boundary detection
is heuristic: any of ``.``, ``!``, ``?``, ``。``, ``!``, ``?`` followed by
whitespace or end-of-string.
"""

from __future__ import annotations

import re
from typing import Any

from docconvert.chunkers.base import BaseChunker
from docconvert.parsers.models import Chunk

_BOUNDARY_RE = re.compile(r"(?<=[.!?。!?])\s+")


class SentenceChunker(BaseChunker):
    """Chunk on sentence boundaries while staying under ``chunk_size``."""

    DEFAULT_CHUNK_SIZE = 512
    DEFAULT_OVERLAP = 64

    def __init__(
        self,
        chunk_size: int = DEFAULT_CHUNK_SIZE,
        chunk_overlap: int = DEFAULT_OVERLAP,
    ):
        if chunk_size <= 0:
            raise ValueError("chunk_size must be > 0")
        if chunk_overlap < 0 or chunk_overlap >= chunk_size:
            raise ValueError("chunk_overlap must be in [0, chunk_size)")
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap

    def chunk(self, content: Any, **kwargs: Any) -> list[Chunk]:
        size = int(kwargs.get("chunk_size", self.chunk_size))
        overlap = int(kwargs.get("chunk_overlap", self.chunk_overlap))
        if size <= 0:
            raise ValueError("chunk_size must be > 0")
        if overlap < 0 or overlap >= size:
            raise ValueError("chunk_overlap must be in [0, chunk_size)")

        doc = self._coerce_to_document(content)
        text = doc.text.strip()
        if not text:
            return []

        sentences = [s for s in _BOUNDARY_RE.split(text) if s]
        if not sentences:
            return []

        chunks: list[str] = []
        current: list[str] = []
        current_len = 0

        for sent in sentences:
            sent_len = len(sent) + (1 if current else 0)
            if current and current_len + sent_len > size:
                chunks.append(" ".join(current))
                overlap_text = " ".join(current)
                if overlap > 0 and len(overlap_text) > overlap:
                    overlap_text = overlap_text[-overlap:]
                current = [overlap_text] if overlap_text else []
                current_len = len(overlap_text)
                if current:
                    current.append(sent)
                    current_len += len(sent) + 1
                else:
                    current.append(sent)
                    current_len = len(sent)
            else:
                current.append(sent)
                current_len += sent_len

        if current:
            tail = " ".join(current)
            if chunks and tail == chunks[-1]:
                pass
            else:
                chunks.append(tail)

        total = len(chunks)
        out: list[Chunk] = []
        cursor = 0
        for i, body in enumerate(chunks):
            start = text.find(body[:40], cursor)
            if start < 0:
                start = cursor
            end = start + len(body)
            cursor = max(end - overlap, start + 1)
            out.append(self._make_chunk(
                text=body,
                metadata=doc.metadata,
                index=i,
                total=total,
                start=start,
                end=end,
            ))
        return out
