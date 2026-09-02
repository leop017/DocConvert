"""Base class for parsers.

A parser converts raw text (Markdown, HTML, plain text) into a structured
:class:`~docconvert.parsers.models.Document`. The chunker layer consumes
that document. Keeping parsers format-specific and chunkers format-agnostic
lets a downstream RAG pipeline mix and match, e.g. ``MarkdownParser`` ->
``SentenceChunker``.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from docconvert.parsers.models import Document


class BaseParser(ABC):
    """Abstract base for all parsers.

    Subclasses must implement :meth:`parse` and accept arbitrary ``kwargs``
    so callers can tune behavior without breaking the contract.
    """

    @abstractmethod
    def parse(self, content: Any, **kwargs: Any) -> Document:
        """Parse ``content`` into a :class:`Document`.

        ``content`` is whatever the subclass accepts. For the built-in
        text parsers that means ``str`` (Markdown / HTML / plain text).
        For subclasses wrapping structured inputs (e.g. an Excel sheet)
        ``content`` may be a richer object.
        """
        ...
