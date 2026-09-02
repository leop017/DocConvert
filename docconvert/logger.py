from __future__ import annotations

import logging
import sys

LOGGER_NAME = "docconvert"


def setup_logging(level: int | str = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)

    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(level)
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(name)s - %(message)s",
            datefmt="%H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    else:
        # A later call (e.g. main.py -> main_cli with --verbose) must
        # raise the already-registered handler's level too; setting only
        # the logger's level is not enough, the handler would keep
        # filtering the messages out.
        for h in logger.handlers:
            h.setLevel(level)

    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger(LOGGER_NAME)
