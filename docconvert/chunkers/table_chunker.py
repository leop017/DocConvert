from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class BaseChunker(ABC):
    @abstractmethod
    def chunk(self, content: Any, **kwargs) -> list[Any]:
        ...
