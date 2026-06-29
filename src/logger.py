"""Centralized logging configuration for email-automation.

Usage:
    from src.logger import get_logger
    log = get_logger(__name__)
    log.info("Hello")
"""

import logging
import sys
from datetime import datetime
from pathlib import Path


def setup_logging(log_dir: Path, level: int = logging.DEBUG) -> None:
    """Configure the root logger once.

    Creates the log directory if needed, adds a timestamped file handler and a
    console handler (INFO+) so that every ``get_logger()`` call inherits both.

    Args:
        log_dir: Directory where log files are written.
        level:   Minimum log level for the file handler (default DEBUG).
    """
    root = logging.getLogger("email_automation")
    if root.handlers:
        # Already set up – avoid duplicate handlers on re-import
        return

    root.setLevel(logging.DEBUG)
    log_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"run_{timestamp}.log"

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler — captures DEBUG and above
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # Console handler — INFO and above
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    root.addHandler(ch)


def get_logger(name: str) -> logging.Logger:
    """Return a child logger under the ``email_automation`` namespace.

    Args:
        name: Typically ``__name__`` of the calling module.

    Returns:
        A ``logging.Logger`` instance inheriting handlers from the root logger.
    """
    return logging.getLogger(f"email_automation.{name}")


__all__ = ["setup_logging", "get_logger"]
