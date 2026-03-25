"""Colored console logging via Rich (level colors + markup in messages)."""

from __future__ import annotations

import logging
import sys


def setup_colored_logging(level: int = logging.INFO) -> None:
    try:
        from rich.console import Console
        from rich.logging import RichHandler
        from rich.theme import Theme
    except ImportError:
        logging.basicConfig(
            level=level,
            format="%(asctime)s | %(levelname)-7s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            stream=sys.stdout,
            force=True,
        )
        logging.getLogger(__name__).warning(
            "Install rich for colors: pip install rich"
        )
        return

    theme = Theme(
        {
            "logging.level.info": "bold cyan",
            "logging.level.warning": "bold yellow",
            "logging.level.error": "bold red",
            "logging.level.critical": "bold white on red",
            "log.time": "dim",
        }
    )
    console = Console(theme=theme, highlight=True)
    handler = RichHandler(
        console=console,
        show_time=True,
        omit_repeated_times=False,
        show_path=False,
        rich_tracebacks=True,
        tracebacks_show_locals=False,
        markup=True,
        log_time_format="[%Y-%m-%d %H:%M:%S]",
    )
    logging.basicConfig(
        level=level,
        format="%(message)s",
        handlers=[handler],
        force=True,
    )
