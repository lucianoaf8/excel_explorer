"""
Exception management utilities
"""

import logging
from typing import Optional

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

class ExplorerError(Exception):
    """Base exception class for Excel Explorer errors."""


def configure_logging(level: int = logging.INFO, log_file: Optional[str] = None) -> None:
    """Configure the root logger.

    Args:
        level: Logging level (e.g. logging.INFO).
        log_file: Optional path to a log file. If omitted, logs go to stderr.
    """
    # Avoid duplicating handlers if configure_logging called multiple times
    if logging.getLogger().handlers:
        return

    handlers = [logging.StreamHandler()]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    logging.basicConfig(level=level, format=LOG_FORMAT, handlers=handlers)


def log_exception(exc: Exception) -> None:
    """Log an exception with traceback."""
    logging.exception("Unhandled exception", exc_info=exc)
