"""Fixed-size character window chunker.

Splits the document into overlapping windows of ``chunk_size`` characters
with ``chunk_overlap`` characters of shared context between neighbours.
This is the lowest-common-denominator strategy: it works on any text but
has no semantic awareness, so a chunk can cut a sentence in half.
"""

from __future__ import annotations

from typing import Any

from docconvert.chunkers.base import BaseChunker
from docconvert.parsers.models import Chunk


class FixedSizeChunker(BaseChunker):
    """Sliding-window chunker measured in characters.

    Parameters
    ----------
    chunk_size:
        Maximum number of characters per chunk. Must be > 0. Defaults to
        ``512`` which is a comfortable size for most embedding models.
    chunk_overlap:
        Number of trailing characters from the previous chunk to prepend
        to the next one. Must be ``0 <= overlap < chunk_size``.
    """

    DEFAULT_CHUNK_SIZE = 512
    DEFAULT_OVERLAP = 64

    def __init__(
        self,
        chunk_size: int = DEFAULT_CHUNK_SIZE,
        chunk_overlap: int = DEFAULT_OVERLAP,
    ):
        if chunk_size <= 0:
            raise ValueError("chunk_size must be > 0")
        if chunk_overlap < 0:
            raise ValueError("chunk_overlap must be >= 0")
        if chunk_overlap >= chunk_size:
            raise ValueError("chunk_overlap must be < chunk_size")
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
        text = doc.text
        if not text:
            return []

        step = size - overlap
        spans: list[tuple[int, int]] = []
        start = 0
        n = len(text)
        while start < n:
            end = min(start + size, n)
            spans.append((start, end))
            if end == n:
                break
            start += step

        total = len(spans)
        return [
            self._make_chunk(
                text=text[s:e],
                metadata=doc.metadata,
                index=i,
                total=total,
                start=s,
                end=e,
            )
            for i, (s, e) in enumerate(spans)
        ]
