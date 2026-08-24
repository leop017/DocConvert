from __future__ import annotations

from typing import Any

from docconvert.exporters.base import BaseExporter


class HtmlExporter(BaseExporter):

    def export(self, data: Any, **kwargs) -> str:
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            return data.get('content', str(data))
        return str(data)
