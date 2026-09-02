"""Plain-text parser.

Treats the input as one large paragraph. Whitespace runs and non-printable
characters are normalized so the resulting ``text`` is safe to feed into a
chunker that expects clean input.
"""

from __future__ import annotations

import re
from typing import Any

from docconvert.parsers.base import BaseParser
from docconvert.parsers.models import Document, Element

_MULTI_WS = re.compile(r"[ \t\f\v]+")
_MULTI_NL = re.compile(r"\n{3,}")


class PlainTextParser(BaseParser):
    """Default fallback parser for non-Markdown / non-HTML inputs."""

    def parse(self, content: Any, **kwargs: Any) -> Document:
        text = self._coerce_to_str(content)
        normalized = _MULTI_WS.sub(" ", text)
        normalized = _MULTI_NL.sub("\n\n", normalized).strip()
        elements = [Element(element_type="paragraph", text=normalized)] if normalized else []
        return Document(text=normalized, elements=elements, metadata={})

    @staticmethod
    def _coerce_to_str(content: Any) -> str:
        if isinstance(content, str):
            return content
        if isinstance(content, bytes):
            return content.decode("utf-8", errors="replace")
        return str(content)
