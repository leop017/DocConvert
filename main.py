#!/usr/bin/env python3
"""
文档转换工具 - Entry Point

Usage:
    python main.py                  # Launch GUI
    python main.py convert <file>   # CLI conversion mode
"""

from __future__ import annotations

import sys

from docconvert.logger import setup_logging


def main():
    setup_logging()
    if len(sys.argv) > 1 and sys.argv[1] == 'convert':
        from docconvert.cli import main_cli
        sys.exit(main_cli(sys.argv[1:]))
    else:
        import tkinter as tk
        from docconvert.gui import DocConvertApp
        root = tk.Tk()
        DocConvertApp(root)
        root.mainloop()


if __name__ == '__main__':
    main()
