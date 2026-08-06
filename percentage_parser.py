# percentage_parser.py
"""
Parser for workbooks containing both 'PERSENTASE %' and 'Summary' sheets.

Sheet layouts (both use same merged-header pattern):
  Row 0: [Tenant, "Jun-26", "",       "Jul-26", "",       "Persentase Kenaikan/Penurunan"]
  Row 1: ["",     "Sales/Month", "Sales/Day", "Sales/Month", "Sales/Day", ""]
  Row 2+: [name,  sales_month,   sales_day,   sales_month,   sales_day,   pct%]
  Row N:  [Total, ...]   ← skip row
  (PERSENTASE % only) Row N+2: [Summary :, ...]
  (PERSENTASE % only) Row N+3+: paragraph text
"""

import re
import calendar as cal
import pandas as pd
from pathlib import Path


# ── Month name → number ───────────────────────────────────────
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
    Parse 'Jun-26', 'Dec-25', 'Jul-2026', 'January-26'.
    Returns (year, month) or None.
    """
    if text is None:
        return None
    s = str(text).strip().lower()
    if not s or s in ("nan", "none", ""):
        return None

    # "Jun-26" / "January-26" / "Jun-2026"
    m = re.fullmatch(r"([a-z]{3,})\s*[-_/]\s*(\d{2,4})", s)
    if m:
        mn = MONTH_NAMES.get(m.group(1))
        yr = int(m.group(2))
        if yr < 100:
            yr += 2000
        if mn and 2020 <= yr <= 2035:
            return (yr, mn)

    # "2026-06"
    m = re.fullmatch(r"(\d{4})\s*[-_/]\s*(\d{1,2})", s)
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
        elif re.fullmatch(r"\d+[.,]\d{1,2}", s):
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
    Parse '5%', '-9%', '0%'.
    Returns float (5.0, -9.0, 0.0) or None.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", ""):
        return None
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*%", s)
    if m:
        return float(m.group(1))
    # Maybe stored as decimal: 0.05 → 5%
    try:
        v = float(s)
        if -1.0 <= v <= 1.0 and v != 0:
            return round(v * 100, 1)
    except ValueError:
        pass
    return None


def _is_skip_row(name):
    """True for total/summary/blank rows that should not be treated as tenants."""
    low = name.lower().strip()
    skip = (
        "total", "grand", "jumlah", "subtotal", "sub total",
        "summary", "ringkasan", "nan", "none", "",
    )
    return any(low == k or low.startswith(k) for k in skip)


def _read_sheet_rows(filepath, sheet_name):
    """Read sheet → list[list[str]]. Returns None on failure."""
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        df = pd.read_excel(
            filepath, sheet_name=sheet_name,
            header=None, engine=engine,
        ).fillna("")
        return [
            [str(c) if str(c) != "" else "" for c in row]
            for row in df.values.tolist()
        ]
    except Exception:
        return None


def _find_sheet(filepath, pattern):
    """
    Case-insensitive fuzzy sheet-name search.
    Returns actual sheet name or None.
    """
    try:
        ext = Path(filepath).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        pat = pattern.lower()
        for name in xl.sheet_names:
            if pat in name.lower():
                return name
        return None
    except Exception:
        return None


def _find_month_header_row(rows, min_months=2):
    """
    Scan rows[:10] for the row that has ≥ min_months month-year cells.
    Returns (header_idx, {col: (year, month)}) or (None, {}).
    """
    for idx, row in enumerate(rows[:10]):
        month_cols = {}
        for c, cell in enumerate(row):
            my = _parse_month_header(cell)
            if my:
                month_cols[c] = my
        if len(month_cols) >= min_months:
            return idx, month_cols
    return None, {}


def _find_pct_col(header_row):
    """Find the Persentase column index in the first header row."""
    for c, cell in enumerate(header_row):
        low = str(cell).strip().lower()
        if any(k in low for k in ("persentase", "percentage", "kenaikan", "penurunan")):
            return c
    return None


def _extract_summary_text(rows):
    """
    Extract paragraph text that follows the Total row and 'Summary :' marker
    in the PERSENTASE % sheet.
    Returns list[str] (one entry per non-empty paragraph).
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

        # Collect all non-empty cells joined into one paragraph string
        text = " ".join(
            str(c).strip() for c in row
            if str(c).strip() and str(c).strip().lower() not in ("nan", "none")
        ).strip()
        if text:
            paragraphs.append(text)

    return paragraphs


# ── Public API ────────────────────────────────────────────────

def detect_percentage_summary_file(filepath):
    """
    Returns (pct_sheet_name, summary_sheet_name) if BOTH sheets exist,
    otherwise None.  Requires BOTH sheets to be present (Q17).
    """
    pct  = _find_sheet(filepath, "persentase")
    summ = _find_sheet(filepath, "summary")
    if pct and summ:
        return pct, summ
    return None


def parse_percentage_summary(filepath):
    """
    Parse a workbook that has both PERSENTASE % and Summary sheets.

    Returns dict:
    {
        "success": True,
        "format": "percentage_summary",
        "is_percentage_summary": True,
        "tenants": {
            "GOGO SUPERMARKET": {
                "monthly":           {"2026-06": 6_285_210_095.5, ...},
                "monthly_daily_avg": {"2026-06": 209_507_003, ...},
                "daily": [],    # always empty — this format has no daily rows
                "files": [],
            },
            ...
        },
        "percentage": {
            "GOGO SUPERMARKET": {"from": "2026-06", "to": "2026-07", "pct": 5.0},
            ...
        },
        "pct_months": {"from": "2026-06", "to": "2026-07"},
        "summary_text": ["paragraph1...", "paragraph2...", "paragraph3..."],
        "all_months":   ["2025-12", "2026-01", ..., "2026-07"],
        "message": "...",
    }
    """
    sheets = detect_percentage_summary_file(filepath)
    if not sheets:
        return None
    pct_sheet, sum_sheet = sheets

    # ═══════════════════════════════════════════════════════════
    # 1. Parse PERSENTASE % sheet
    # ═══════════════════════════════════════════════════════════
    pct_rows = _read_sheet_rows(filepath, pct_sheet)
    if not pct_rows or len(pct_rows) < 3:
        return None

    hdr_idx, month_cols = _find_month_header_row(pct_rows)
    if hdr_idx is None or len(month_cols) < 2:
        return None

    # Always exactly 2 months in PERSENTASE sheet (Q1)
    sorted_mc   = sorted(month_cols.items())          # [(col, (yr,mn)), ...]
    m1_col, m1_ym = sorted_mc[0]
    m2_col, m2_ym = sorted_mc[1]
    m1_key = f"{m1_ym[0]}-{m1_ym[1]:02d}"
    m2_key = f"{m2_ym[0]}-{m2_ym[1]:02d}"

    pct_col = _find_pct_col(pct_rows[hdr_idx])

    # Tenant data dictionary
    tenants   = {}
    pct_data  = {}

    # Data starts at hdr_idx + 2 (skip both header rows)
    for row in pct_rows[hdr_idx + 2:]:
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
        pv  = _parse_pct(row[pct_col] if pct_col is not None and pct_col < len(row) else None)

        if sm1 is None and sm2 is None:
            continue

        monthly     = {}
        daily_avg   = {}
        if sm1 is not None:
            monthly[m1_key]   = sm1
        if sd1 is not None:
            daily_avg[m1_key] = sd1
        if sm2 is not None:
            monthly[m2_key]   = sm2
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
            # Columns come in pairs: Sales/Month at col c, Sales/Day at col c+1
            sorted_smc = sorted(smonth_cols.items())   # [(col, (yr,mn)), ...]

            for row in sum_rows[shdr_idx + 2:]:
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
                    key = f"{yr}-{mn:02d}"

                    sm = _parse_number(row[col_idx]     if col_idx     < len(row) else None)
                    sd = _parse_number(row[col_idx + 1] if col_idx + 1 < len(row) else None)

                    # Summary sheet is canonical for historical data
                    if sm is not None:
                        tenants[name]["monthly"][key] = sm
                    if sd is not None:
                        tenants[name]["monthly_daily_avg"][key] = sd

    # ═══════════════════════════════════════════════════════════
    # 3. Collate all months across both sheets
    # ═══════════════════════════════════════════════════════════
    all_months = sorted({
        k
        for t in tenants.values()
        for k in t.get("monthly", {})
    })

    if not all_months:
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
            f"Percentage + Summary: {len(tenants)} tenant(s) × {len(all_months)} month(s) "
            f"({all_months[0]} → {all_months[-1]}). "
            f"MoM: {m1_key} vs {m2_key}."
        ),
    }
