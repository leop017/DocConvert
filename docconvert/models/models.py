from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class MergeInfo:
    rowspan: int = 1
    colspan: int = 1
    is_master: bool = False
    is_merged: bool = False
    min_row: int = 0
    min_col: int = 0
    max_row: int = 0
    max_col: int = 0


@dataclass
class ProgressEvent:
    message: str = ""
    progress: float = 0.0
    done: bool = False
    error: Optional[str] = None
