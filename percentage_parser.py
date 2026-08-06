# percentage_parser.py
"""
Parser for workbooks containing both 'PERSENTASE %' and 'Summary' sheets.

PERSENTASE % sheet layout:
  Row 0: [Tenant, "Jun-26", "",       "Jul-26", "",       "Persentase Kenaikan/Penurunan"]
  Row 1: ["",     "Sales/Month","Sales/Day","Sales/Month","Sales/Day",""]
  Row 2+: [tenant_name, sales_month_1, sales_day_1, sales_month_2, sales_day_2, pct]
  ...
  Row N: [Total, total_1, "", total_2, "", ""]

Summary sheet layout:
  Row 0: [Tenant, "Dec-25","",       "Jan-26","",       "Feb-26","", ...]
  Row 1: ["",     "Sales/Month","Sales/Day","Sales/Month","Sales/Day",...]
  Row 2+: [tenant_name, sales_month, sales_day, sales_month, sales_day, ...]
  Row N:  [Total, ...]
"""

import re
import calendar as cal
import pandas as pd
from pathlib import Path

# ── Month name mapping (shared with app.py) ──────────────────

MONTH_NAMES = {
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
    "nov": 11, "november": 11, "nop": 11, "nopember": 11,
    "dec": 12, "december": 12, "des": 12, "desember": 12,
}


def _parse_month_col_header(text):
    """
    Parse column headers like 'Jun-26', 'Dec-25', 'Jul-2026', 'January-26'.
    Returns (year, month) or None.
    """
    if text is None:
        return None
    s = str(text).strip().lower()
    if not s or s in ("nan", "none", ""):
        return None

    # Pattern: "Jun-26", "Dec-25", "Jul-2026", "January-26"
    m = re.fullmatch(r"([a-z]{3,})\s*[-_/\.]\s*(\d{2,4})", s)
    if m:
        mn = MONTH_NAMES.get(m.group(1))
        yr = int(m.group(2))
        if yr < 100:
            yr += 2000
        if mn and 2020 <= yr <= 2035:
            return (yr, mn)

    # Pattern: "2026-06"
    m = re.fullmatch(r"(\d{4})\s*[-_/\.]\s*(\d{1,2})", s)
    if m:
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            return (yr, mn)

    return None


def _parse_number(raw):
    """Parse number from cell value (handles Indonesian formatting)."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", "n/a", ""):
        return None
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s or s in ("-", "--"):
        return None
    neg = bool(re.match(r"^-\d", s))
    s = s.lstrip("-")
    if not s:
        return None
    if re.fullmatch(r"\d{1,3}([.,]\d{3})+", s):
        val = float(s.replace(".", "").replace(",", ""))
    elif re.fullmatch(r"\d+", s):
        val = float(s)
    elif re.fullmatch(r"\d+[.,]\d{1,2}", s):
        val = float(s.replace(",", "."))
    else:
        digits = re.sub(r"[^\d]", "", s)
        if not digits:
            return None
        val = float(digits)
    return -val if neg else val


def _parse_pct(raw):
    """Parse percentage like '5%', '-9%', '0%'. Returns float like 5.0, -9.0."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", ""):
        return None
    m = re.match(r"(-?\d+(?:\.\d+)?)\s*%?", s.replace(",", "."))
    if m:
        return float(m.group(1))
    return None


def _is_skip_row(name):
    """Check if a row should be skipped (total, summary, blank)."""
    low = name.lower().strip()
    skip_kw = (
        "total", "grand", "jumlah", "subtotal", "sub total",
        "summary", "ringkasan", "keterangan", "catatan",
        "nan", "none", "",
    )
    return any(k == low or low.startswith(k) for k in skip_kw)


def _read_sheet(filepath, sheet_name):
    """Read a specific sheet from an Excel file, return list of lists."""
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine=engine)
        df = df.fillna("")
        rows = []
        for row in df.values.tolist():
            rows.append([str(c) if str(c) != "" else "" for c in row])
        return rows
    except Exception:
        return None


def _find_sheet_name(filepath, pattern):
    """
    Find a sheet name matching a pattern (case-insensitive fuzzy).
    Returns the actual sheet name or None.
    """
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        sheet_names = xl.sheet_names

        for name in sheet_names:
            if pattern.lower() in name.lower():
                return name
        return None
    except Exception:
        return None


def _parse_month_columns(header_row):
    """
    Parse the first header row to find month columns.
    Returns dict: { col_index: (year, month) }

    Layout: [Tenant, "Jun-26", "", "Jul-26", "", "Persentase..."]
    Month header appears at the Sales/Month position, next col is Sales/Day.
    """
    month_cols = {}
    for c, cell in enumerate(header_row):
        my = _parse_month_col_header(cell)
        if my:
            month_cols[c] = my
    return month_cols


def _extract_summary_text(rows, start_after_total=True):
    """
    Extract summary paragraph text from PERSENTASE % sheet.
    Looks for text after the Total row and "Summary :" marker.
    Returns list of paragraph strings.
    """
    paragraphs = []
    found_total = False
    found_summary_marker = False

    for row in rows:
        first_cell = str(row[0]).strip() if row else ""
        low = first_cell.lower()

        if not found_total:
            if low.startswith("total") or low.startswith("grand total"):
                found_total = True
            continue

        # After total row — look for "Summary" marker
        if not found_summary_marker:
            if "summary" in low or "ringkasan" in low:
                found_summary_marker = True
            continue

        # Collect paragraph text
        # Join all non-empty cells in the row
        text = " ".join(str(c).strip() for c in row if str(c).strip() and str(c).strip() != "nan")
        text = text.strip()
        if text:
            paragraphs.append(text)

    return paragraphs


def detect_percentage_summary_file(filepath):
    """
    Check if an Excel file contains both required sheets.
    Returns (pct_sheet_name, summary_sheet_name) or None.
    """
    pct_name = _find_sheet_name(filepath, "persentase")
    summary_name = _find_sheet_name(filepath, "summary")

    if pct_name and summary_name:
        return (pct_name, summary_name)
    return None


def parse_percentage_summary(filepath):
    """
    Full parser for workbooks with PERSENTASE % and Summary sheets.

    Returns:
    {
        "success": True,
        "format": "percentage_summary",
        "is_percentage_summary": True,
        "tenants": {
            "GOGO SUPERMARKET": {
                "monthly": {"2026-06": 6285210095.5, "2026-07": 6596418164.0, ...},
                "monthly_daily_avg": {"2026-06": 209507003, "2026-07": 212787683, ...},
                "daily": [],
                "files": [],
            }
        },
        "percentage": {
            "GOGO SUPERMARKET": {"from": "2026-06", "to": "2026-07", "pct": 5.0},
        },
        "pct_months": {"from": "2026-06", "to": "2026-07"},
        "summary_text": ["paragraph1...", "paragraph2..."],
        "all_months": ["2025-12", "2026-01", ...],
        "message": "..."
    }
    """
    sheets = detect_percentage_summary_file(filepath)
    if not sheets:
        return None

    pct_sheet_name, summary_sheet_name = sheets

    # ═══════════════════════════════════════════════════════════
    # Parse PERSENTASE % sheet
    # ═══════════════════════════════════════════════════════════
    pct_rows = _read_sheet(filepath, pct_sheet_name)
    if not pct_rows or len(pct_rows) < 3:
        return None

    # Find header rows (row with month names)
    pct_month_cols = {}
    header_row_idx = None
    for idx, row in enumerate(pct_rows[:10]):
        mc = _parse_month_columns(row)
        if len(mc) >= 2:
            pct_month_cols = mc
            header_row_idx = idx
            break

    if header_row_idx is None or len(pct_month_cols) < 2:
        return None

    # Identify month pairs: each month has Sales/Month (col) and Sales/Day (col+1)
    sorted_month_cols = sorted(pct_month_cols.items(), key=lambda x: x[0])
    month_1_col, month_1_ym = sorted_month_cols[0]
    month_2_col, month_2_ym = sorted_month_cols[1]

    month_1_key = f"{month_1_ym[0]}-{month_1_ym[1]:02d}"
    month_2_key = f"{month_2_ym[0]}-{month_2_ym[1]:02d}"

    # Find percentage column
    pct_col = None
    for c, cell in enumerate(pct_rows[header_row_idx]):
        low = str(cell).strip().lower()
        if "persentase" in low or "percentage" in low or "kenaikan" in low:
            pct_col = c
            break

    # Parse tenant data from PERSENTASE % sheet
    pct_tenants = {}
    pct_data = {}

    for row in pct_rows[header_row_idx + 2:]:  # skip both header rows
        if not row:
            continue
        name = str(row[0]).strip()
        if not name or _is_skip_row(name):
            continue
        if not re.search(r"[a-zA-Z]", name):
            continue

        # Sales/Month values
        sm1 = _parse_number(row[month_1_col] if month_1_col < len(row) else None)
        sd1 = _parse_number(row[month_1_col + 1] if month_1_col + 1 < len(row) else None)
        sm2 = _parse_number(row[month_2_col] if month_2_col < len(row) else None)
        sd2 = _parse_number(row[month_2_col + 1] if month_2_col + 1 < len(row) else None)

        pct_val = None
        if pct_col is not None and pct_col < len(row):
            pct_val = _parse_pct(row[pct_col])

        if sm1 is None and sm2 is None:
            continue

        monthly = {}
        monthly_daily_avg = {}

        if sm1 is not None:
            monthly[month_1_key] = sm1
        if sd1 is not None:
            monthly_daily_avg[month_1_key] = sd1
        if sm2 is not None:
            monthly[month_2_key] = sm2
        if sd2 is not None:
            monthly_daily_avg[month_2_key] = sd2

        pct_tenants[name] = {
            "monthly": monthly,
            "monthly_daily_avg": monthly_daily_avg,
            "daily": [],
            "files": [],
        }

        if pct_val is not None:
            pct_data[name] = {
                "from": month_1_key,
                "to": month_2_key,
                "pct": pct_val,
            }

    # Extract summary text
    summary_text = _extract_summary_text(pct_rows)

    # ═══════════════════════════════════════════════════════════
    # Parse Summary sheet (historical months)
    # ═══════════════════════════════════════════════════════════
    sum_rows = _read_sheet(filepath, summary_sheet_name)
    if not sum_rows or len(sum_rows) < 3:
        # If summary sheet fails, still return PERSENTASE data
        return {
            "success": True,
            "format": "percentage_summary",
            "is_percentage_summary": True,
            "tenants": pct_tenants,
            "percentage": pct_data,
            "pct_months": {"from": month_1_key, "to": month_2_key},
            "summary_text": summary_text,
            "all_months": sorted(set(
                k for t in pct_tenants.values() for k in t["monthly"]
            )),
            "message": f"Percentage sheet: {len(pct_tenants)} tenants, "
                       f"{month_1_key} vs {month_2_key}. "
                       f"Summary sheet could not be read.",
        }

    # Find header row in summary sheet
    sum_month_cols = {}
    sum_header_idx = None
    for idx, row in enumerate(sum_rows[:10]):
        mc = _parse_month_columns(row)
        if len(mc) >= 2:
            sum_month_cols = mc
            sum_header_idx = idx
            break

    if sum_header_idx is None:
        # Return just percentage data
        return {
            "success": True,
            "format": "percentage_summary",
            "is_percentage_summary": True,
            "tenants": pct_tenants,
            "percentage": pct_data,
            "pct_months": {"from": month_1_key, "to": month_2_key},
            "summary_text": summary_text,
            "all_months": sorted(set(
                k for t in pct_tenants.values() for k in t["monthly"]
            )),
            "message": f"Percentage sheet: {len(pct_tenants)} tenants, "
                       f"{month_1_key} vs {month_2_key}. "
                       f"Summary sheet header not detected.",
        }

    # Build month mapping: col_idx -> (year, month, key)
    # Each month takes 2 columns: Sales/Month, Sales/Day
    sorted_sum_months = sorted(sum_month_cols.items(), key=lambda x: x[0])

    # Merge summary data into tenants
    all_tenants = dict(pct_tenants)  # start with pct data

    for row in sum_rows[sum_header_idx + 2:]:  # skip both header rows
        if not row:
            continue
        name = str(row[0]).strip()
        if not name or _is_skip_row(name):
            continue
        if not re.search(r"[a-zA-Z]", name):
            continue

        if name not in all_tenants:
            all_tenants[name] = {
                "monthly": {},
                "monthly_daily_avg": {},
                "daily": [],
                "files": [],
            }

        for col_idx, (yr, mn) in sorted_sum_months:
            key = f"{yr}-{mn:02d}"

            # Sales/Month
            sm = _parse_number(row[col_idx] if col_idx < len(row) else None)
            if sm is not None:
                # Summary sheet provides the canonical historical data
                all_tenants[name]["monthly"][key] = sm

            # Sales/Day (next column)
            sd = _parse_number(row[col_idx + 1] if col_idx + 1 < len(row) else None)
            if sd is not None:
                if "monthly_daily_avg" not in all_tenants[name]:
                    all_tenants[name]["monthly_daily_avg"] = {}
                all_tenants[name]["monthly_daily_avg"][key] = sd

    # Collect all months
    all_months = sorted(set(
        k for t in all_tenants.values() for k in t.get("monthly", {})
    ))

    tenant_count = len(all_tenants)
    month_count = len(all_months)

    return {
        "success": True,
        "format": "percentage_summary",
        "is_percentage_summary": True,
        "tenants": all_tenants,
        "percentage": pct_data,
        "pct_months": {"from": month_1_key, "to": month_2_key},
        "summary_text": summary_text,
        "all_months": all_months,
        "message": (
            f"Percentage + Summary: {tenant_count} tenants × {month_count} months "
            f"({all_months[0]} to {all_months[-1]}). "
            f"MoM comparison: {month_1_key} vs {month_2_key}."
        ),
    }
