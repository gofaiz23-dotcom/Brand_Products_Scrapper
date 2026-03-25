"""Configuration for Homelegance base-sheet → scraper output tool."""

from pathlib import Path

TOOL_DIR = Path(__file__).resolve().parent
HOMETOOL_DIR = TOOL_DIR.parent

# --- Excel paths ---
# Put your master SKU workbook here (any of the names below, or the only .xlsx in folder).
BASE_SHEETS_DIR = HOMETOOL_DIR.parent / "sheets" / "Homelagance" / "basesheets"


def get_base_sheet_path() -> Path:
    """Resolve base workbook: preferred names first, else newest .xlsx in BASE_SHEETS_DIR."""
    preferred = [
        BASE_SHEETS_DIR / "Homelegance-All-CategoriesSkus-InSingleSheet (1).xlsx",
        BASE_SHEETS_DIR / "Homelegance-All-CategoriesSkus-InSingleSheet.xlsx",
    ]
    for p in preferred:
        if p.is_file():
            return p.resolve()
    if BASE_SHEETS_DIR.is_dir():
        all_xlsx = list(BASE_SHEETS_DIR.glob("*.xlsx"))
        if all_xlsx:
            return max(all_xlsx, key=lambda x: x.stat().st_mtime).resolve()
    return preferred[0]
# Live “success only” xlsx while scraping (open this file to watch rows appear).
LIVE_SUCCESS_XLSX_DIR = (
    HOMETOOL_DIR.parent / "sheets" / "Homelagance" / "scrppedSheets"
)
# Resolves to: .../ProductsScrapper/sheets/Homelagance/scrppedSheets
# Per run (same stamp across singles + multiples). strftime pattern: dd-mm-yy_HHMMSS
# (slashes not used — invalid in Windows paths.)
RUN_FILE_STAMP_FORMAT = "%d-%m-%y_%H%M%S"
#   Homelagance-{stamp}-single-subskus-sheets.xlsx
#   Homelagance-{stamp}-multiple-subskus-from-singles.xlsx  (all Sub-SKUs found in singles)
#   Homelagance-{stamp}-multiple-subskus-sheets.xlsx        (site scrape by master SKU)

# Multi rows resolved only from singles file: these base columns + attributes (JSON).
MULTI_NARROW_BASE_HEADERS: list[str] = [
    "Brand Name",
    "Category",
    "Collection Name",
    "Product Links",
    "Single / Set Item",
    "Ship Type",
    "NEW/Master SKU",
    "Sub-SKU",
    "Comments",
]

# Column header in the base sheet (exact match)
SUB_SKU_HEADER = "Sub-SKU"
MASTER_SKU_HEADER = "NEW/Master SKU"

# --- Playwright ---
# False = visible browser window; True = run in background (no window).
HEADLESS = False
# Milliseconds between Playwright actions (0 = off). Use e.g. 250 when debugging.
SLOW_MO_MS = 0

# Used to resolve relative image / link paths from HTML
SITE_ORIGIN = "https://shopla.homelegance.com"

# Seconds to wait after each product (be gentle on the server)
DELAY_BETWEEN_SKUS_SEC = 1.5

# Navigation / element timeouts (ms)
DEFAULT_TIMEOUT_MS = 45_000

# None = process every eligible row; set to a small int for smoke tests
MAX_SKUS = None
