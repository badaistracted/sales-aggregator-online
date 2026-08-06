# percentage_parser.py
"""
Parser for workbooks containing both 'PERSENTASE %' and 'Summary' sheets.

Actual Excel layout (as read by pandas with header=None):
  Row 0: [Tenant, "2026-06-01 00:00:00", "", "2026-07-01 00:00:00", "", "Persentase Kenaikan/Penurunan"]
  Row 1: ["",     "Sales/Month", "Sales/Day", "Sales/Month", "Sales/Day", ""]
  Row 2+: [name, sales_month, sales_day, sales_month, sales_day, pct_decimal]
  Row N:  [Total, ...]
  Row N+2: [Summary :, ...]
  Row N+3+: paragraph text

Summary sheet layout:
  Row 0: [Tenant, "2025-12-01 00:00:00", "", "2026-01-01 00:00:00", "", ...]
  Row 1: ["",     "Sales/Month", "Sales/Day", "Sales/Month", "Sales/Day", ...]
  Row 2+: [name, sales_month, sales_day, ...]
  Row N:  [Total, ...]
"""

import re
import calendar as cal
import pandas as pd
from pathlib import Path


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


def _parse_month_header(text):
    """
    Parse month-year from a cell. Handles ALL of these formats:
      "2026-06-01 00:00:00"   <- actual format from pandas datetime
      "2026-06-01"
      "2026-06"
      "Jun-26"
      "June-2026"
      "jun-26"
    Returns (year, month) tuple or None.
    """
    if text is None:
        return None
    s = str(text).strip()
    if not s or s.lower() in ("nan", "none", "", "-", "--"):
        return None

    # ── Format 1: "2026-06-01 00:00:00" or "2026-06-01" (pandas datetime) ──
    m = re.match(r"^(\d{4})-(\d{1,2})-\d{1,2}", s)
    if m:
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            return (yr, mn)

    sl = s.lower()

    # ── Format 2: "Jun-26", "jun-2026", "June-26" ──
    m = re.fullmatch(r"([a-z]{3,})\s*[-_/\.]\s*(\d{2,4})", sl)
    if m:
        mn = MONTH_NAMES.get(m.group(1))
        yr = int(m.group(2))
        if yr < 100:
            yr += 2000
        if mn and 2020 <= yr <= 2035:
            return (yr, mn)

    # ── Format 3: "2026-06" ──
    m = re.fullmatch(r"(\d{4})\s*[-_/\.]\s*(\d{1,2})", sl)
    if m:
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            return (yr, mn)

    return None


def _parse_number(raw):
    """Parse number handling Indonesian thousand separators."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", "n/a", ""):
        return None

    # Try float directly first (handles "6285210095.5" perfectly)
    try:
        v = float(s)
        # Sanity check: reject if it looks like a percentage decimal
        # (those are handled by _parse_pct, not here)
        return v
    except ValueError:
        pass

    # Indonesian formatting: remove separators
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s or s in ("-", "--"):
        return None
    neg = s.startswith("-")
    s = s.lstrip("-")
    if not s:
        return None

    try:
        if re.fullmatch(r"\d{1,3}([.,]\d{3})+", s):
            val = float(s.replace(".", "").replace(",", ""))
        elif re.fullmatch(r"\d+", s):
            val = float(s)
        elif re.fullmatch(r"\d+[.,]\d+", s):
            val = float(s.replace(",", "."))
        else:
            digits = re.sub(r"[^\d]", "", s)
            val = float(digits) if digits else None
    except (ValueError, TypeError):
        return None

    if val is None:
        return None
    return -val if neg else val


def _parse_pct(raw):
    """
    Parse percentage. Handles ALL of these:
      0.04951434618276562   -> 5.0   (stored as decimal fraction)
      -0.0903051438278627   -> -9.0
      "5%"                  -> 5.0
      "-9%"                 -> -9.0
      "0%"                  -> 0.0
    Returns float (percentage points) or None.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", ""):
        return None

    # "5%", "-9%", "44%"
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*%", s)
    if m:
        return float(m.group(1))

    # Raw decimal: 0.049... or -0.090... stored as fraction
    try:
        v = float(s)
        # It's a fraction (between -2 and 2, excluding 0)
        # Convert to percentage points
        if -2.0 <= v <= 2.0:
            return round(v * 100, 1)
    except ValueError:
        pass

    return None


def _is_skip_row(name):
    """True for total/summary/blank rows."""
    low = name.lower().strip()
    skip = (
        "total", "grand", "jumlah", "subtotal", "sub total",
        "summary", "ringkasan", "nan", "none", "",
    )
    return any(low == k or low.startswith(k) for k in skip)


def _read_sheet_rows(filepath, sheet_name):
    """Read sheet → list[list[str]]."""
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        df = pd.read_excel(
            filepath, sheet_name=sheet_name,
            header=None, engine=engine,
        ).fillna("")
        rows = []
        for row in df.values.tolist():
            rows.append([str(c) if str(c) != "" else "" for c in row])
        return rows
    except Exception as e:
        print("[PCT PARSER] Error reading sheet '%s': %s" % (sheet_name, e))
        return None


def _find_sheet(filepath, pattern):
    """Case-insensitive fuzzy sheet-name search."""
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        pat = pattern.lower()
        for name in xl.sheet_names:
            if pat in name.lower():
                return name
        return None
    except Exception as e:
        print("[PCT PARSER] Error listing sheets: %s" % e)
        return None


def _find_month_header_row(rows, min_months=2):
    """
    Scan rows[:15] for the row that has >= min_months month-year cells.
    Returns (header_idx, {col: (year, month)}) or (None, {}).
    """
    for idx, row in enumerate(rows[:15]):
        month_cols = {}
        for c, cell in enumerate(row):
            my = _parse_month_header(cell)
            if my:
                month_cols[c] = my
        if len(month_cols) >= min_months:
            return idx, month_cols
    return None, {}


def _find_subheader_row(rows, hdr_idx):
    """
    Check if the row after the month-header row is a subheader
    (contains 'sales', 'month', 'day' etc.).
    Returns data_start_idx.
    """
    sub_idx = hdr_idx + 1
    if sub_idx >= len(rows):
        return sub_idx

    sub_check = " ".join(str(c).lower() for c in rows[sub_idx])
    if any(k in sub_check for k in ("sales", "month", "day", "penjualan")):
        return hdr_idx + 2

    return hdr_idx + 1


def _find_pct_col(header_row):
    """Find the percentage column index."""
    for c, cell in enumerate(header_row):
        low = str(cell).strip().lower()
        if any(k in low for k in ("persentase", "percentage", "kenaikan", "penurunan")):
            return c
    return None


def _extract_summary_text(rows):
    """
    Extract paragraph text after 'Total' row and 'Summary :' marker.
    """
    passed_total = False
    passed_marker = False
    paragraphs = []

    for row in rows:
        first = str(row[0]).strip().lower() if row else ""

        if not passed_total:
            if first.startswith("total") or first.startswith("grand total"):
                passed_total = True
            continue

        if not passed_marker:
            if "summary" in first or "ringkasan" in first:
                passed_marker = True
            continue

        # Collect non-empty cells
        text = " ".join(
            str(c).strip() for c in row
            if str(c).strip()
            and str(c).strip().lower() not in ("nan", "none", "-", "--")
        ).strip()
        if text:
            paragraphs.append(text)

    return paragraphs


# ── Public API ────────────────────────────────────────────────

def detect_percentage_summary_file(filepath):
    """
    Returns (pct_sheet_name, summary_sheet_name) if BOTH sheets exist.
    Requires BOTH sheets — rejects if only one found.
    """
    pct = _find_sheet(filepath, "persentase")
    summ = _find_sheet(filepath, "summary")
    if pct and summ:
        return pct, summ
    return None


def parse_percentage_summary(filepath):
    """
    Parse a workbook that has both PERSENTASE % and Summary sheets.

    Returns the standard result dict or None on failure.
    """
    sheets = detect_percentage_summary_file(filepath)
    if not sheets:
        return None
    pct_sheet, sum_sheet = sheets

    # ═══════════════════════════════════════════════════════════
    # 1. Parse PERSENTASE sheet
    # ═══════════════════════════════════════════════════════════
    pct_rows = _read_sheet_rows(filepath, pct_sheet)
    if not pct_rows or len(pct_rows) < 3:
        return None

    hdr_idx, month_cols = _find_month_header_row(pct_rows)
    if hdr_idx is None or len(month_cols) < 2:
        return None

    sorted_mc = sorted(month_cols.items())
    m1_col, m1_ym = sorted_mc[0]
    m2_col, m2_ym = sorted_mc[1]
    m1_key = "%d-%02d" % (m1_ym[0], m1_ym[1])
    m2_key = "%d-%02d" % (m2_ym[0], m2_ym[1])

    pct_col = _find_pct_col(pct_rows[hdr_idx])
    data_start = _find_subheader_row(pct_rows, hdr_idx)

    tenants = {}
    pct_data = {}

    for row in pct_rows[data_start:]:
        if not row:
            continue
        name = str(row[0]).strip()
        if not name or _is_skip_row(name):
            continue
        if not re.search(r"[a-zA-Z]", name):
            continue

        sm1 = _parse_number(row[m1_col]     if m1_col     < len(row) else None)
        sd1 = _parse_number(row[m1_col + 1] if m1_col + 1 < len(row) else None)
        sm2 = _parse_number(row[m2_col]     if m2_col     < len(row) else None)
        sd2 = _parse_number(row[m2_col + 1] if m2_col + 1 < len(row) else None)
        pv  = _parse_pct(
            row[pct_col] if pct_col is not None and pct_col < len(row) else None
        )

        if sm1 is None and sm2 is None:
            continue

        monthly = {}
        daily_avg = {}
        if sm1 is not None:
            monthly[m1_key] = sm1
        if sd1 is not None:
            daily_avg[m1_key] = sd1
        if sm2 is not None:
            monthly[m2_key] = sm2
        if sd2 is not None:
            daily_avg[m2_key] = sd2

        tenants[name] = {
            "monthly":           monthly,
            "monthly_daily_avg": daily_avg,
            "daily":             [],
            "files":             [],
        }
        if pv is not None:
            pct_data[name] = {"from": m1_key, "to": m2_key, "pct": pv}

    summary_text = _extract_summary_text(pct_rows)

    # ═══════════════════════════════════════════════════════════
    # 2. Parse Summary sheet (historical months)
    # ═══════════════════════════════════════════════════════════
    sum_rows = _read_sheet_rows(filepath, sum_sheet)

    if sum_rows and len(sum_rows) >= 3:
        shdr_idx, smonth_cols = _find_month_header_row(sum_rows, min_months=2)

        if shdr_idx is not None:
            sorted_smc = sorted(smonth_cols.items())
            sum_data_start = _find_subheader_row(sum_rows, shdr_idx)

            for row in sum_rows[sum_data_start:]:
                if not row:
                    continue
                name = str(row[0]).strip()
                if not name or _is_skip_row(name):
                    continue
                if not re.search(r"[a-zA-Z]", name):
                    continue

                if name not in tenants:
                    tenants[name] = {
                        "monthly":           {},
                        "monthly_daily_avg": {},
                        "daily":             [],
                        "files":             [],
                    }

                for col_idx, (yr, mn) in sorted_smc:
                    key = "%d-%02d" % (yr, mn)
                    sm = _parse_number(row[col_idx]     if col_idx     < len(row) else None)
                    sd = _parse_number(row[col_idx + 1] if col_idx + 1 < len(row) else None)
                    if sm is not None:
                        tenants[name]["monthly"][key] = sm
                    if sd is not None:
                        tenants[name]["monthly_daily_avg"][key] = sd

    # ═══════════════════════════════════════════════════════════
    # 3. Collate all months
    # ═══════════════════════════════════════════════════════════
    all_months = sorted(set(
        k
        for t in tenants.values()
        for k in t.get("monthly", {})
    ))

    if not tenants:
        return None

    return {
        "success":               True,
        "format":                "percentage_summary",
        "is_percentage_summary": True,
        "tenants":               tenants,
        "percentage":            pct_data,
        "pct_months":            {"from": m1_key, "to": m2_key},
        "summary_text":          summary_text,
        "all_months":            all_months,
        "message": (
            "Percentage + Summary: %d tenant(s) x %d month(s) (%s to %s). MoM: %s vs %s."
            % (
                len(tenants),
                len(all_months),
                all_months[0] if all_months else "?",
                all_months[-1] if all_months else "?",
                m1_key,
                m2_key,
            )
        ),
    }
