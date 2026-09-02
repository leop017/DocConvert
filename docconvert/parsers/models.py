"""Shared data classes for parsers and chunkers.

The :class:`Document` produced by a parser is the single source of truth that
chunker implementations consume. Keeping these dataclasses here avoids a
circular dependency between ``docconvert.parsers`` and ``docconvert.chunkers``,
and keeps the existing ``docconvert.models`` module (which holds controller /
runner-level dataclasses) cleanly separated from the RAG-pipeline primitives.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class Element:
    """A single block-level unit recovered from the source document.

    ``element_type`` is one of: ``heading``, ``paragraph``, ``list``,
    ``code``, ``table``, ``image``, ``other``. The string is intentionally
    loose (no enum) so third-party parsers can introduce new types without
    breaking the public contract.
    """

    element_type: str = "other"
    text: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class Document:
    """A parsed document, the common currency between parsers and chunkers.

    ``text`` is the flattened plain-text view (always safe to use as a
    fallback). ``elements`` preserves structural information when the parser
    could recover it. ``metadata`` carries format-specific hints (e.g. the
    ``<title>`` of an HTML page, the YAML front-matter of a Markdown file).
    """

    text: str = ""
    elements: list[Element] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class Chunk:
    """A contiguous slice of a :class:`Document`, ready for embedding.

    ``text`` is what downstream consumers feed to an embedding model;
    ``metadata`` is copied (shallow) from the parent and may be augmented
    with chunk-local hints such as ``chunk_index`` and ``chunk_count``.
    """

    text: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)
    start_index: int = 0
    end_index: int = 0
