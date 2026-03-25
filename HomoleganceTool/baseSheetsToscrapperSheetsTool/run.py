"""
Entry point: singles scraper, then multi Sub-SKU pass (split outputs).

Outputs under sheets/Homelagance/scrppedSheets (same stamp dd-mm-yy_HHMMSS):
  Homelagance-{stamp}-single-subskus-sheets.xlsx
  Homelagance-{stamp}-multiple-subskus-from-singles.xlsx  (optional)
  Homelagance-{stamp}-multiple-subskus-sheets.xlsx        (optional)

Run (from this folder):

    python run.py
"""

from __future__ import annotations

import logging

from log_theme import setup_colored_logging

setup_colored_logging()

from multiplesNewMaterskuNaserSubskuFindings import run as run_multiples  # noqa: E402
from singleSUB_SKUscraper import run as run_singles  # noqa: E402


if __name__ == "__main__":
    singles_path, singles_rows, singles_att, singles_plan = run_singles()
    narrow_path, web_path, narrow_rows, web_rows, web_ok = run_multiples(singles_path)

    total_rows = singles_rows + narrow_rows + web_rows

    logging.info(
        "[bold white on green] ═══ RUN FINISHED ═══ [/] "
        "[dim]total rows written across all output files:[/] [bold cyan]%d[/]",
        total_rows,
    )
    logging.info(
        "[green]Singles[/]: [bold]%d[/] rows → [cyan]%s[/] "
        "[dim](attempts %d / planned %d)[/]",
        singles_rows,
        singles_path,
        singles_att,
        singles_plan,
    )
    if narrow_rows and narrow_path:
        logging.info(
            "[green]Multiples (from singles only)[/]: [bold]%d[/] rows → [cyan]%s[/]",
            narrow_rows,
            narrow_path,
        )
    if web_path:
        logging.info(
            "[green]Multiples (site scrape)[/]: [bold]%d[/] rows "
            "([green]%d[/] ok, [red]%d[/] failed) → [cyan]%s[/]",
            web_rows,
            web_ok,
            web_rows - web_ok,
            web_path,
        )
