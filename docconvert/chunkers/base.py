"""Base class for chunkers.

A chunker takes a parsed :class:`~docconvert.parsers.models.Document` (or
a raw string) and slices it into :class:`Chunk` objects that downstream
embedding / retrieval code can consume independently.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from docconvert.parsers.models import Chunk, Document


class BaseChunker(ABC):
    """Abstract base for chunkers."""

    @abstractmethod
    def chunk(self, content: Any, **kwargs: Any) -> list[Chunk]:
        """Split ``content`` into a list of :class:`Chunk`.

        ``content`` may be either a :class:`Document` (preferred) or a raw
        ``str`` (in which case the chunker is responsible for any
        normalization it needs).
        """
        ...

    @staticmethod
    def _coerce_to_document(content: Any) -> Document:
        if isinstance(content, Document):
            return content
        if isinstance(content, str):
            return Document(text=content)
        return Document(text=str(content))

    @staticmethod
    def _make_chunk(
        text: str,
        metadata: dict[str, Any],
        *,
        index: int,
        total: int,
        start: int,
        end: int,
    ) -> Chunk:
        meta = dict(metadata)
        meta["chunk_index"] = index
        meta["chunk_count"] = total
        return Chunk(text=text, metadata=meta, start_index=start, end_index=end)
