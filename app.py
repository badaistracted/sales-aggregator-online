# app.py
"""
Tenant Sales Aggregator — Handles pivot-style Excel files where:
  - Column headers are month-year codes like: Dec-25, Jan-26, Feb-26
  - Row labels are tenant names
  - Cell values are sales figures
Also still handles the traditional Date/Sales column format.
"""

import os
import re
import uuid
import warnings
import calendar as cal
from datetime import date, timedelta, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import holidays
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ══════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

UPLOAD_FOLDER = Path("uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)
SESSION_STORE: dict[str, dict] = {}
ALLOWED_EXT = {".xlsx", ".xls"}

def allowed_file(fn: str) -> bool:
    return Path(fn).suffix.lower() in ALLOWED_EXT


# ══════════════════════════════════════════════════════════════════
#  SECTION 1 — Indonesian Calendar
# ══════════════════════════════════════════════════════════════════

def get_id_holidays(year: int) -> set:
    return set(holidays.Indonesia(years=year).keys())

def classify_day(d: date, hols: set) -> str:
    if d.weekday() >= 5:
        return "Weekend"
    if d in hols:
        return "Weekend"
    return "Weekday"

def build_timeline(year: int, month: int) -> pd.DataFrame:
    hols = get_id_holidays(year)
    first = date(year, month, 1)
    n = cal.monthrange(year, month)[1]
    rows = []
    for i in range(n):
        d = first + timedelta(days=i)
        rows.append({
            "Date": d,
            "DayName": d.strftime("%A"),
            "DayType": classify_day(d, hols),
        })
    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════
#  SECTION 2 — Month-Year Header Parser
# ══════════════════════════════════════════════════════════════════

# All formats people use for month-year headers
MONTH_MAP = {
    "jan": 1, "january": 1, "januari": 1,
    "feb": 2, "february": 2, "februari": 2,
    "mar": 3, "march": 3, "maret": 3,
    "apr": 4, "april": 4,
    "may": 5, "mei": 5,
    "jun": 6, "june": 6, "juni": 6,
    "jul": 7, "july": 7, "juli": 7,
    "aug": 8, "august": 8, "agustus": 8, "agu": 8, "ags": 8,
    "sep": 9, "september": 9, "sept": 9,
    "oct": 10, "october": 10, "okt": 10, "oktober": 10,
    "nov": 11, "november": 11, "nop": 11,
    "dec": 12, "december": 12, "des": 12, "desember": 12,
}


def parse_month_year(text: str) -> tuple[int, int] | None:
    """
    Parse strings like: Dec-25, Jan-26, Feb 2026, 2025-12,
    December 2025, Des-25, etc.
    Returns (year_4digit, month_number) or None.
    """
    if text is None:
        return None
    s = str(text).strip().lower()
    if not s:
        return None

    # ── Pattern 1: "Mon-YY" or "Mon YY" or "Mon-YYYY" ────────────────
    m1 = re.match(r"^([a-z]+)[.\-_/\s]+(\d{2,4})$", s)
    if m1:
        month_str, year_str = m1.group(1), m1.group(2)
        month_num = MONTH_MAP.get(month_str)
        if month_num:
            yr = int(year_str)
            if yr < 100:
                yr += 2000
            return (yr, month_num)

    # ── Pattern 2: "YYYY-MM" or "YY-Mon" ─────────────────────────────
    m2 = re.match(r"^(\d{2,4})[.\-_/\s]+([a-z]+|\d{1,2})$", s)
    if m2:
        year_str, month_str = m2.group(1), m2.group(2)
        yr = int(year_str)
        if yr < 100:
            yr += 2000
        # month might be a number or name
        if month_str.isdigit():
            mn = int(month_str)
            if 1 <= mn <= 12:
                return (yr, mn)
        else:
            mn = MONTH_MAP.get(month_str)
            if mn:
                return (yr, mn)

    # ── Pattern 3: "December 2025" or "2025 December" ─────────────────
    m3 = re.match(r"^([a-z]+)\s+(\d{4})$", s)
    if m3:
        mn = MONTH_MAP.get(m3.group(1))
        if mn:
            return (int(m3.group(2)), mn)

    m4 = re.match(r"^(\d{4})\s+([a-z]+)$", s)
    if m4:
        mn = MONTH_MAP.get(m4.group(2))
        if mn:
            return (int(m4.group(1)), mn)

    return None


def is_month_header(val) -> bool:
    """Quick check if a value looks like a month-year header."""
    return parse_month_year(str(val)) is not None


# ══════════════════════════════════════════════════════════════════
#  SECTION 3 — Pivot Format Scanner
# ══════════════════════════════════════════════════════════════════

class ScanResult:
    def __init__(self):
        self.success      = False
        self.format       = ""      # "pivot" or "columnar"
        self.tenants_data = {}      # {tenant_name: {(year,month): sales_value}}
        self.header_row   = -1
        self.label_col    = -1
        self.warnings     = []
        self.error        = ""


def _clean_number(val) -> float:
    """Convert messy number strings to float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0.0
        return float(val)
    s = str(val).strip()
    if not s or s.lower() in ("nan", "-", "n/a", ""):
        return 0.0
    # Remove currency symbols and thousand separators
    cleaned = re.sub(r"[Rp,\s$€£]", "", s)
    # Handle Indonesian dot-as-thousands (e.g., 5.000.000)
    # If there are multiple dots, they're thousands separators
    if cleaned.count(".") > 1:
        cleaned = cleaned.replace(".", "")
    elif cleaned.count(".") == 1:
        # Could be decimal or thousands — if 3 digits after dot, it's thousands
        parts = cleaned.split(".")
        if len(parts[1]) == 3:
            cleaned = cleaned.replace(".", "")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def smart_scan(filepath: Path) -> ScanResult:
    """
    Scan an Excel file to detect its format:

    FORMAT 1 — "Pivot" (your format):
        Row headers = tenant names
        Column headers = month-year codes (Dec-25, Jan-26, etc.)
        Cell values = monthly sales figures

    FORMAT 2 — "Columnar" (traditional):
        Column: Date | Sales
        Row: one row per day per tenant
    """
    result = ScanResult()
    ext    = filepath.suffix.lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"

    # ── Read raw (no header assumption) ───────────────────────────────
    try:
        raw = pd.read_excel(filepath, header=None, dtype=str, engine=engine)
    except Exception as exc:
        result.error = f"Cannot open file: {exc}"
        return result

    if raw.empty:
        result.error = "File is empty."
        return result

    # ── Scan every row to find the one with month-year headers ────────
    max_scan = min(50, len(raw))
    best_row = -1
    best_month_count = 0
    best_month_cols = {}      # col_index → (year, month)

    for row_idx in range(max_scan):
        row_vals = raw.iloc[row_idx].tolist()
        month_cols = {}
        for col_idx, val in enumerate(row_vals):
            parsed = parse_month_year(str(val) if val is not None else "")
            if parsed:
                month_cols[col_idx] = parsed

        if len(month_cols) > best_month_count:
            best_month_count = len(month_cols)
            best_row = row_idx
            best_month_cols = month_cols

    # ── Decide format ─────────────────────────────────────────────────
    if best_month_count >= 2:
        # PIVOT FORMAT detected
        result.format     = "pivot"
        result.header_row = best_row
        result = _parse_pivot(raw, result, best_row, best_month_cols)
    else:
        # Try columnar format as fallback
        result.format = "columnar"
        result = _parse_columnar(filepath, raw, result, engine)

    return result


def _parse_pivot(
    raw: pd.DataFrame,
    result: ScanResult,
    header_row: int,
    month_cols: dict,       # {col_index: (year, month)}
) -> ScanResult:
    """
    Parse the pivot format where:
    - header_row contains month-year codes
    - Rows below contain: [tenant_name, ..., sales, sales, sales, ...]
    """
    result.header_row = header_row

    # Figure out which column has the tenant labels
    # It's typically the first non-month column, or column 0
    header_vals = raw.iloc[header_row].tolist()
    month_col_indices = set(month_cols.keys())

    # Find the label column (first column that is NOT a month header)
    label_col = None
    for c_idx in range(len(header_vals)):
        if c_idx not in month_col_indices:
            val = str(header_vals[c_idx]).strip().lower() if header_vals[c_idx] is not None else ""
            # Skip if it's empty
            if val and val not in ("nan", ""):
                label_col = c_idx
                break

    # If no explicit label column found, use column 0
    if label_col is None:
        label_col = 0
        # But if column 0 IS a month column, try to find another
        if label_col in month_col_indices:
            for c in range(len(header_vals)):
                if c not in month_col_indices:
                    label_col = c
                    break

    result.label_col = label_col

    # ── Extract tenant data from rows below the header ────────────────
    tenants_data = {}
    data_start = header_row + 1

    for row_idx in range(data_start, len(raw)):
        row_vals = raw.iloc[row_idx].tolist()

        # Get tenant name from the label column
        tenant_name = str(row_vals[label_col]).strip() if label_col < len(row_vals) else ""

        # Skip empty rows, total rows, header-like rows
        if not tenant_name or tenant_name.lower() in (
            "", "nan", "total", "grand total", "jumlah", "subtotal",
            "sub total", "sum", "rata-rata", "average", "rata rata",
        ):
            continue

        # Skip if the tenant name looks like a number (probably a data value in wrong col)
        try:
            float(tenant_name.replace(",", "").replace(".", ""))
            continue   # it's a number, not a tenant name
        except ValueError:
            pass

        # Extract monthly sales for this tenant
        monthly_sales = {}
        for col_idx, (yr, mn) in month_cols.items():
            if col_idx < len(row_vals):
                val = _clean_number(row_vals[col_idx])
                monthly_sales[(yr, mn)] = val

        if monthly_sales:
            tenants_data[tenant_name] = monthly_sales

    if not tenants_data:
        result.error = "Found month-year headers but no tenant data rows."
        return result

    result.tenants_data = tenants_data
    result.success = True

    # Report what was found
    months_found = sorted(set(
        ym for td in tenants_data.values() for ym in td.keys()
    ))
    
