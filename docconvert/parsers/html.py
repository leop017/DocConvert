"""HTML parser.

Uses :mod:`beautifulsoup4` (already a project dependency) to extract the
visible text of a page along with its top-level block structure. Script
and style tags are dropped. The result is suitable for downstream chunking.
"""

from __future__ import annotations

from typing import Any

from bs4 import BeautifulSoup, Tag

from docconvert.parsers.base import BaseParser
from docconvert.parsers.models import Document, Element

_DROP_TAGS = {"script", "style", "noscript", "iframe"}
_BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "pre", "blockquote", "tr", "div"}


class HtmlParser(BaseParser):
    """A small BeautifulSoup-based HTML parser."""

    def parse(self, content: Any, **kwargs: Any) -> Document:
        html = self._coerce_to_str(content)
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup.find_all(_DROP_TAGS):
            tag.decompose()

        metadata: dict[str, Any] = {}
        title_tag = soup.find("title")
        if isinstance(title_tag, Tag):
            title_text = title_tag.get_text(strip=True)
            if title_text:
                metadata["title"] = title_text

        elements: list[Element] = []
        body_node: Tag = soup.find("body") or soup  # type: ignore[assignment]
        if not isinstance(body_node, Tag):
            body_node = soup
        for child in list(body_node.children):
            if not isinstance(child, Tag):
                continue
            self._collect(child, elements)

        flat = "\n\n".join(e.text for e in elements if e.text)
        return Document(text=flat, elements=elements, metadata=metadata)

    @staticmethod
    def _coerce_to_str(content: Any) -> str:
        if isinstance(content, str):
            return content
        if isinstance(content, bytes):
            return content.decode("utf-8", errors="replace")
        return str(content)

    def _collect(self, node: Tag, out: list[Element]) -> None:
        name = (node.name or "").lower()
        if name in _BLOCK_TAGS or name in {"ul", "ol", "table"}:
            text = node.get_text("\n", strip=True)
            if text:
                if name.startswith("h") and len(name) == 2 and name[1].isdigit():
                    out.append(Element(
                        element_type="heading",
                        text=text,
                        metadata={"level": int(name[1])},
                    ))
                elif name in {"ul", "ol"}:
                    out.append(Element(element_type="list", text=text))
                elif name == "table":
                    rows: list[str] = []
                    for tr in node.children:
                        if not isinstance(tr, Tag) or tr.name != "tr":
                            continue
                        cells = [c for c in tr.children
                                 if isinstance(c, Tag) and c.name in {"td", "th"}]
                        rows.append(" | ".join(c.get_text(" ", strip=True) for c in cells))
                    out.append(Element(element_type="table", text=text, metadata={"rows": rows}))
                elif name == "pre":
                    out.append(Element(element_type="code", text=text))
                else:
                    out.append(Element(element_type="paragraph", text=text))
            for child in node.children:
                if isinstance(child, Tag) and child.name not in _BLOCK_TAGS \
                        and child.name not in {"ul", "ol", "table"}:
                    self._collect(child, out)
        else:
            for child in node.children:
                if isinstance(child, Tag):
                    self._collect(child, out)
