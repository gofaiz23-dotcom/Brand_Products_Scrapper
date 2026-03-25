"""
Entry point: configures logging, then runs the single-Sub-SKU Homelegance scraper.

Run command (from this folder: baseSheetsToscrapperSheetsTool):

    python run.py

Example (PowerShell):

    cd C:\\Users\\HI\\Desktop\\ProductsScrapper\\HomoleganceTool\\baseSheetsToscrapperSheetsTool
    python run.py
"""

from __future__ import annotations

import logging

from log_theme import setup_colored_logging

setup_colored_logging()

from singleSUB_SKUscraper import run  # noqa: E402


if __name__ == "__main__":
    out = run()
    logging.info("[bold green]Run complete.[/] Output: [cyan]%s[/cyan]", out)
