from docconvert.exporters.base import BaseExporter
from docconvert.exporters.html import HtmlExporter
from docconvert.exporters.markdown import MarkdownExporter
from docconvert.exporters.json_exporter import JsonExporter

__all__ = [
    "BaseExporter",
    "HtmlExporter",
    "MarkdownExporter",
    "JsonExporter",
    "get_exporter",
]


def get_exporter(fmt: str) -> BaseExporter:
    if fmt == 'html':
        return HtmlExporter()
    elif fmt == 'md':
        return MarkdownExporter()
    elif fmt == 'json':
        return JsonExporter()
    raise ValueError(f'不支持的输出格式： {fmt}')
