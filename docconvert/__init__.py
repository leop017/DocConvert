"""DocConvert - 文档转换工具

The GUI is imported lazily so CLI / library use (``python main.py convert``
or ``from docconvert.controller import ConversionController``) does not
require the Tkinter package.
"""

__all__ = ["DocConvertApp"]


def __getattr__(name: str):
    if name == "DocConvertApp":
        from docconvert.gui.app import DocConvertApp
        return DocConvertApp
    raise AttributeError(f"module 'docconvert' has no attribute {name!r}")
