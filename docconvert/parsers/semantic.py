"""Backwards-compatible re-export.

Historically this module hosted :class:`BaseParser`. The class moved to
:mod:`docconvert.parsers.base` so that the parser package can grow with
format-specific implementations (``markdown`` / ``html`` / ``plain``).
External callers that still do ``from docconvert.parsers.semantic import
BaseParser`` keep working via this shim.
"""

from docconvert.parsers.base import BaseParser

__all__ = ["BaseParser"]
