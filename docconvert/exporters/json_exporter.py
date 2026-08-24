from __future__ import annotations

from typing import Any

from docconvert.exporters.base import BaseExporter


class JsonExporter(BaseExporter):

    def export(self, data: Any, **kwargs) -> str:
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            import json
            return json.dumps(data, ensure_ascii=False, indent=2)
        return str(data)
