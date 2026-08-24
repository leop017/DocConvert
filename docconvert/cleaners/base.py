from __future__ import annotations

from abc import ABC, abstractmethod


class BaseCleaner(ABC):

    @abstractmethod
    def clean(self, content: str, **kwargs) -> str:
        ...
