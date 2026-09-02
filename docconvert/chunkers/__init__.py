"""Chunkers — split parsed documents into embedding-sized slices.

The submodule mirrors :mod:`docconvert.parsers`: one ABC plus a handful
of format-aware and format-agnostic implementations, and a factory that
resolves a name to an instance.
"""

from docconvert.chunkers.base import BaseChunker
from docconvert.chunkers.fixed_size import FixedSizeChunker
from docconvert.chunkers.markdown import MarkdownChunker
from docconvert.chunkers.sentence import SentenceChunker


def get_chunker(strategy: str, **kwargs) -> BaseChunker:
    """Factory mirroring :func:`docconvert.parsers.get_parser`.

    Accepted ``strategy`` values: ``"fixed"`` / ``"fixed_size"``,
    ``"sentence"``, ``"markdown"`` / ``"md"``. Anything else raises
    :class:`ValueError`. Extra ``kwargs`` are forwarded to the
    chunker constructor (so callers can configure ``chunk_size`` /
    ``chunk_overlap``).
    """
    key = strategy.lower()
    if key in {"fixed", "fixed_size"}:
        return FixedSizeChunker(**kwargs)
    if key == "sentence":
        return SentenceChunker(**kwargs)
    if key in {"markdown", "md"}:
        return MarkdownChunker(**kwargs)
    raise ValueError(f"Unsupported chunking strategy: {strategy!r}")


__all__ = [
    "BaseChunker",
    "FixedSizeChunker",
    "MarkdownChunker",
    "SentenceChunker",
    "get_chunker",
]
