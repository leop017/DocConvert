from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, Optional

from docconvert.config import AppConfig, DEFAULT_CONFIG
from docconvert.logger import get_logger


class BaseConverter(ABC):

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG
        self.logger = get_logger()
        self.cancel_check: Optional[Callable[[], bool]] = None

    @abstractmethod
    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        ...

    def _export(self, content: str, output_fmt: str) -> str:
        from docconvert.exporters import get_exporter
        exporter = get_exporter(output_fmt)
        return exporter.export(content)
