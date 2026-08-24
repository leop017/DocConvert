from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class AppConfig:
    chunk_size: int = 1000
    max_rows: int = 0
    markdown_style: str = "github"
    preview_chars: int = 5000
    preview_lines: int = 150
    large_file_size: int = 20 * 1024 * 1024

    cleaning_rules: dict[str, bool] = field(default_factory=lambda: {
        "remove_page_numbers": True,
        "remove_duplicate_headers": True,
        "remove_empty_lines": True,
        "normalize_spaces": True,
    })


DEFAULT_CONFIG = AppConfig()
