"""Parsers — turn raw text into structured :class:`Document` objects.

The submodule tree is small on purpose: the public surface is ``BaseParser``
plus the three built-in implementations, and every concrete parser returns
the same :class:`~docconvert.parsers.models.Document` shape so chunkers can
swap freely.
"""

from docconvert.parsers.base import BaseParser
from docconvert.parsers.html import HtmlParser
from docconvert.parsers.markdown import MarkdownParser
from docconvert.parsers.models import Chunk, Document, Element
from docconvert.parsers.plain import PlainTextParser


def get_parser(fmt: str) -> BaseParser:
    """Factory mirroring :func:`docconvert.exporters.get_exporter`.

    Accepted ``fmt`` values: ``"markdown"`` / ``"md"``, ``"html"`` /
    ``"htm"``, ``"text"`` / ``"txt"`` / ``"plain"``. Anything else raises
    :class:`ValueError`.
    """
    key = fmt.lower().lstrip(".")
    if key in {"markdown", "md"}:
        return MarkdownParser()
    if key in {"html", "htm"}:
        return HtmlParser()
    if key in {"text", "txt", "plain"}:
        return PlainTextParser()
    raise ValueError(f"Unsupported parser format: {fmt!r}")


__all__ = [
    "BaseParser",
    "Chunk",
    "Document",
    "Element",
    "HtmlParser",
    "MarkdownParser",
    "PlainTextParser",
    "get_parser",
]
