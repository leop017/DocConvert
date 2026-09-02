from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Optional

from docconvert.config import DEFAULT_CONFIG, AppConfig


class BaseExporter(ABC):

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG

    @abstractmethod
    def export(self, data: Any, **kwargs) -> str:
        ...
