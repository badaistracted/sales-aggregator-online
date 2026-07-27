# app.py
# ─────────────────────────────────────────────────────────────────
# Tenant Sales Aggregator — Full Stack Single File
# Deploy on Railway: just needs this file + requirements.txt
# ─────────────────────────────────────────────────────────────────

import os
import re
import uuid
import json
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
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024   # 100 MB

UPLOAD_FOLDER = Path("uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)

SESSION_STORE: dict[str, dict] = {}

ALLOWED_EXT = {".xlsx", ".xls"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


# ══════════════════════════════════════════════════════════════════
#  SECTION 1 — Indonesian Calendar
# ══════════════════════════════════════════════════════════════════

def get_indonesian_holidays(year: int) -> set:
    return set(holidays.Indonesia(years=year).keys())


def classify_day(d: date, holiday_set: set) -> str:
    if d.weekday() >= 5:
        return "Weekend"
    if d in holiday_set:
        return "Weekend"
    return "Weekday"


def build_continuous_timeline(year: int, month: int) -> pd.DataFrame:
    holiday_set = get_indonesian_holidays(year)
    first_day   = date(year, month, 1)
    n_days      = cal.monthrange(year, month)[1]
    records     = []
    for i in range(n_days):
        d = first_day + timedelta(days=i)
        records.append({
            "Date"    : d,
            "DayName" : d.strftime("%A"),
            "DayType" : classify_day(d, holiday_set),
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
#  SECTION 2 — Smart File Scanner
# ══════════════════════════════════════════════════════════════════

DATE_ALIASES = [
    "date", "tanggal", "tgl", "hari", "periode", "period",
    "transaction date", "trans date", "trx date", "trx_date",
    "trans_date", "transaction_date", "waktu", "datetime",
    "tgl transaksi", "tanggal transaksi", "tanggal transaksi",
]

SALES_ALIASES = [
    "sales", "penjualan", "total", "amount", "revenue",
    "total sales", "total penjualan", "jumlah", "nilai",
    "net sales", "gross sales", "omzet", "omset",
    "total amount", "sales amount", "income", "pendapatan",
    "total revenue", "daily sales", "nilai penjualan",
    "gross revenue", "net revenue", "turnover",
]


def _normalise(text: str) -> str:
    return re.sub(r"\s+", " ", str(text).lower().strip())


def _find_column(candidates: list, aliases: list) -> str | None:
    norm_map = {_normalise(c): c for c in candidates}

    # Exact match
    for alias in aliases:
        if alias in norm_map:
            return norm_map[alias]

    # Starts with
    for alias in aliases:
        for norm, original in norm_map.items():
            if norm.startswith(alias):
                return original

    # Contains
    for alias in aliases:
        for norm, original in norm_map.items():
            if alias in norm:
                return original

    # Reverse contains
    for alias in aliases:
        for norm, original in norm_map.items():
            if norm and norm in alias:
                return original

    return None


def _looks_like_date(value) -> bool:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return False
    if isinstance(value, date):
        return True
    if hasattr(value, "date"):
        return True
    if isinstance(value, (int, float)):
        if 30000 < value < 60000:
            return True
    if isinstance(value, str):
        patterns = [
            r"\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}",
            r"\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}",
            r"\d{1,2}\s+\w+\s+\d{4}",
        ]
        for p in patterns:
            if re.search(p, str(value).strip()):
                return True
    return False


def _looks_like_number(value) -> bool:
    if value is None:
        return False
    try:
        if isinstance(value, float) and np.isnan(value):
            return False
    except Exception:
        pass
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        cleaned = re.sub(r"[Rp,\.\s$€£]", "", str(value)).strip()
        try:
            float(cleaned)
            return True
        except ValueError:
            return False
    return False


def _score_row_as_header(row_values: list) -> float:
    non_null = [v for v in row_values
                if v is not None and str(v).strip() not in ("", "nan")]
    if not non_null:
        return 0.0

    score      = 0.0
    text_count = 0

    for v in non_null:
        s    = str(v).strip()
        norm = _normalise(s)

        if isinstance(v, str) and len(s) < 60:
            text_count += 1

        if norm in DATE_ALIASES or any(a in norm for a in DATE_ALIASES):
            score += 2.0
        if norm in SALES_ALIASES or any(a in norm for a in SALES_ALIASES):
            score += 2.0

        if _looks_like_date(v):
            score -= 1.5
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            score -= 0.5

    if non_null:
        score += (text_count / len(non_null)) * 1.5

    return score


def _sniff_date_column(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        sample    = df[col].dropna().head(10).tolist()
        date_hits = sum(1 for v in sample if _looks_like_date(v))
        if date_hits >= max(1, len(sample) * 0.5):
            return str(col)
    return None


def _sniff_numeric_column(df: pd.DataFrame, exclude: str | None = None) -> str | None:
    best_col  = None
    best_mean = 0
    for col in df.columns:
        if str(col) == str(exclude):
            continue
        try:
            cleaned = (
                df[col].astype(str)
                .str.replace(r"[Rp,\.\s$€£]", "", regex=True)
                .str.strip()
            )
            numeric = pd.to_numeric(cleaned, errors="coerce").dropna()
            if len(numeric) == 0:
                continue
            m = numeric.mean()
            if m > best_mean:
                best_mean = m
                best_col  = str(col)
        except Exception:
            continue
    return best_col


class ScanResult:
    def __init__(self):
        self.success    : bool             = False
        self.header_row : int              = -1
        self.date_col   : str              = ""
        self.sales_col  : str              = ""
        self.df         : pd.DataFrame | None = None
        self.warnings   : list[str]        = []
        self.error      : str              = ""


def smart_scan(filepath: Path, year: int, month: int) -> ScanResult:
    result = ScanResult()
    ext    = filepath.suffix.lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"

    # ── Raw read (no header assumption) ──────────────────────────────
    try:
        raw = pd.read_excel(filepath, header=None, dtype=str, engine=engine)
    except Exception as exc:
        result.error = f"Cannot open file: {exc}"
        return result

    if raw.empty:
        result.error = "File is empty."
        return result

    # ── Score every row ───────────────────────────────────────────────
    max_scan = min(50, len(raw))
    scores   = []
    for i in range(max_scan):
        row_vals = raw.iloc[i].tolist()
        scores.append((_score_row_as_header(row_vals), i))

    scores.sort(reverse=True)
    best_score, best_row = scores[0]

    # ── Fallback if no clear header ───────────────────────────────────
    if best_score < 0.5:
        result.warnings.append(
            f"No clear header row found (score {best_score:.2f}). "
            "Trying content-based detection."
        )
        header_row = -1
        for i in range(max_scan):
            row = raw.iloc[i].tolist()
            if any(_looks_like_date(v) for v in row) and \
               any(_looks_like_number(v) for v in row):
                header_row = max(0, i - 1)
                break
        if header_row == -1:
            result.error = (
                "Could not find a header row. "
                "Make sure your file has a Date column and a Sales column."
            )
            return result
    else:
        header_row = best_row

    result.header_row = header_row

    # ── Re-read with proper header ────────────────────────────────────
    try:
        df = pd.read_excel(filepath, header=header_row, engine=engine)
    except Exception as exc:
        result.error = f"Failed to re-read with header at row {header_row}: {exc}"
        return result

    df.dropna(axis=1, how="all", inplace=True)
    df.dropna(axis=0, how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        result.error = "No data found after the header row."
        return result

    # ── Fuzzy column matching ─────────────────────────────────────────
    col_names = [str(c) for c in df.columns]
    date_col  = _find_column(col_names, DATE_ALIASES)
    sales_col = _find_column(col_names, SALES_ALIASES)

    if not date_col:
        date_col = _sniff_date_column(df)
        if date_col:
            result.warnings.append(f"Date column detected by content: '{date_col}'")

    if not sales_col:
        sales_col = _sniff_numeric_column(df, exclude=date_col)
        if sales_col:
            result.warnings.append(f"Sales column detected by content: '{sales_col}'")

    if not date_col:
        result.error = (
            f"No Date column found. Columns in file: {col_names}. "
            "Expected names like: date, tanggal, tgl, transaction date, etc."
        )
        return result

    if not sales_col:
        result.error = (
            f"No Sales column found. Columns in file: {col_names}. "
            "Expected names like: sales, penjualan, total, amount, revenue, etc."
        )
        return result

    result.date_col  = date_col
    result.sales_col = sales_col

    # ── Parse and clean ───────────────────────────────────────────────
    work = df[[date_col, sales_col]].copy()
    work.columns = ["Date", "Sales"]

    work["Date"] = pd.to_datetime(work["Date"], errors="coerce", dayfirst=True)
    bad_dates    = work["Date"].isna().sum()
    if bad_dates:
        result.warnings.append(f"{bad_dates} rows had unreadable dates and were skipped.")
    work.dropna(subset=["Date"], inplace=True)
    work["Date"] = work["Date"].dt.date

    # Filter to target month
    work = work[
        work["Date"].apply(lambda d: d.year  == year) &
        work["Date"].apply(lambda d: d.month == month)
    ]

    if work.empty:
        result.error = (
            f"No data found for {month}/{year}. "
            "Check that the file contains data for the selected month and year."
        )
        return result

    # Clean sales numbers
    work["Sales"] = (
        work["Sales"].astype(str)
        .str.replace(r"[Rp,\.\s$€£]", "", regex=True)
        .str.strip()
    )
    work["Sales"] = pd.to_numeric(work["Sales"], errors="coerce").fillna(0)
    work["Sales"] = work["Sales"].clip(lower=0)

    # Handle duplicate dates
    dupes = work.duplicated(subset=["Date"], keep=False).sum()
    if dupes:
        result.warnings.append(
            f"{dupes} duplicate dates found — keeping the highest sales value per date."
        )
        work = (
            work.sort_values("Sales", ascending=False)
                .drop_duplicates(subset=["Date"], keep="first")
        )

    work.sort_values("Date", inplace=True)
    work.reset_index(drop=True, inplace=True)

    result.df      = work
    result.success = True
    return result


# ══════════════════════════════════════════════════════════════════
#  SECTION 3 — Multi-Tenant Aggregator
# ══════════════════════════════════════════════════════════════════

def parse_all_tenants(
    file_paths: list[Path],
    year: int,
    month: int,
) -> tuple[dict[str, pd.DataFrame], list[dict]]:
    """
    Returns:
        tenants_data : { stem: merged_df }
        scan_reports : list of per-file scan detail dicts
    """
    timeline     = build_continuous_timeline(year, month)
    tenants_data = {}
    scan_reports = []

    for fp in file_paths:
        fp     = Path(fp)
        result = smart_scan(fp, year, month)

        report = {
            "file"       : fp.name,
            "success"    : result.success,
            "header_row" : result.header_row + 1 if result.success else None,
            "date_col"   : result.date_col,
            "sales_col"  : result.sales_col,
            "warnings"   : result.warnings,
            "error"      : result.error,
        }
        scan_reports.append(report)

        if result.success:
            merged         = timeline.merge(result.df, on="Date", how="left")
            merged["Sales"] = merged["Sales"].fillna(0)
            tenants_data[fp.stem] = merged

    return tenants_data, scan_reports


# ══════════════════════════════════════════════════════════════════
#  SECTION 4 — Excel Export (Styles)
# ══════════════════════════════════════════════════════════════════

def _fill(hex_col: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_col)

def _font(hex_col: str, bold: bool = False, size: int = 11) -> Font:
    return Font(color=hex_col, bold=bold, size=size, name="Calibri")

THIN = Border(
    left=Side(style="thin", color="BDBDBD"), right=Side(style="thin", color="BDBDBD"),
    top=Side(style="thin", color="BDBDBD"),  bottom=Side(style="thin", color="BDBDBD"),
)
MED = Border(
    left=Side(style="medium", color="7F8C8D"), right=Side(style="medium", color="7F8C8D"),
    top=Side(style="medium", color="7F8C8D"),  bottom=Side(style="medium", color="7F8C8D"),
)

HDR_BG   = "2E4057"
HDR_FG   = "FFFFFF"
WE_BG    = "D5F5E3"
WE_FG    = "1D6A39"
ALT_BG   = "EBF5FB"
TOT_BG   = "D6EAF8"
TOT_FG   = "1B4F72"
ACC_BG   = "2874A6"
TEN_BG   = "1A5276"


def _style_header(ws, row: int, n_cols: int) -> None:
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill      = _fill(HDR_BG)
        cell.font      = _font(HDR_FG, bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = THIN


def _style_data(ws, row: int, n_cols: int, bg: str, fg: str,
                num_cols: set | None = None) -> None:
    num_cols = num_cols or set()
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill   = _fill(bg)
        cell.font   = _font(fg)
        cell.border = THIN
        if c in num_cols:
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")


def _set_widths(ws, widths: dict) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


# ── Sheet builders ────────────────────────────────────────────────

def _build_summary(wb, tenants_data, year, month):
    ws = wb.create_sheet("Monthly_Summary")
    ws.sheet_view.showGridLines = False

    month_label = date(year, month, 1).strftime("%B %Y")
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value     = f"Monthly Sales Summary — {month_label}"
    t.fill      = _fill(ACC_BG)
    t.font      = Font(color="FFFFFF", bold=True, size=16, name="Calibri")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 36

    hdrs = ["Tenant Name", "Total Monthly Sales (IDR)",
            "Weekday Avg/Day (IDR)", "Weekend/Holiday Avg/Day (IDR)",
            "Weekend vs Weekday Ratio"]
    for c, h in enumerate(hdrs, 1):
        ws.cell(row=2, column=c, value=h)
    _style_header(ws, 2, len(hdrs))
    _set_widths(ws, {1: 22, 2: 28, 3: 28, 4: 32, 5: 26})

    grand = 0
    for i, (name, df) in enumerate(tenants_data.items()):
        row    = 3 + i
        bg     = ALT_BG if i % 2 == 0 else "FFFFFF"
        total  = df["Sales"].sum()
        wd_avg = df.loc[df["DayType"] == "Weekday", "Sales"].mean()
        we_avg = df.loc[df["DayType"] == "Weekend", "Sales"].mean()
        ratio  = round(we_avg / wd_avg, 2) if wd_avg and not pd.isna(wd_avg) else 0
        grand += total

        ws.cell(row=row, column=1, value=name.replace("_", " "))
        ws.cell(row=row, column=2, value=total)
        ws.cell(row=row, column=3, value=round(wd_avg, 0) if not pd.isna(wd_avg) else 0)
        ws.cell(row=row, column=4, value=round(we_avg, 0) if not pd.isna(we_avg) else 0)
        ws.cell(row=row, column=5, value=ratio)
        _style_data(ws, row, len(hdrs), bg, "2C3E50", num_cols={2, 3, 4})
        ws.cell(row=row, column=5).alignment = Alignment(horizontal="center")

    tr = 3 + len(tenants_data)
    ws.cell(row=tr, column=1, value="GRAND TOTAL")
    ws.cell(row=tr, column=2, value=grand)
    for c in range(3, len(hdrs) + 1):
        ws.cell(row=tr, column=c, value="—")
    for c in range(1, len(hdrs) + 1):
        cell = ws.cell(row=tr, column=c)
        cell.fill      = _fill(TOT_BG)
        cell.font      = _font(TOT_FG, bold=True, size=12)
        cell.border    = MED
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tr, column=2).number_format = "#,##0"
    ws.cell(row=tr, column=2).alignment = Alignment(horizontal="right")


def _build_tenant_sheet(wb, name, df, year, month):
    safe = (name[:31])
    ws   = wb.create_sheet(safe)
    ws.sheet_view.showGridLines = False

    month_label = date(year, month, 1).strftime("%B %Y")
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value     = f"{name.replace('_', ' ')} — Daily Sales ({month_label})"
    t.fill      = _fill(TEN_BG)
    t.font      = Font(color="FFFFFF", bold=True, size=14, name="Calibri")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    hdrs = ["Date", "Day Name", "Day Type", "Sales (IDR)", "Notes"]
    for c, h in enumerate(hdrs, 1):
        ws.cell(row=2, column=c, value=h)
    _style_header(ws, 2, len(hdrs))
    _set_widths(ws, {1: 14, 2: 14, 3: 14, 4: 20, 5: 20})

    for i, (_, row) in enumerate(df.iterrows()):
        r    = 3 + i
        is_w = row["DayType"] == "Weekend"
        bg   = WE_BG if is_w else "FFFFFF"
        fg   = WE_FG if is_w else "2C3E50"
        note = ("Holiday" if is_w and row["DayName"] not in
                ("Saturday", "Sunday") else "")

        ws.cell(row=r, column=1, value=row["Date"])
        ws.cell(row=r, column=2, value=row["DayName"])
        ws.cell(row=r, column=3, value=row["DayType"])
        ws.cell(row=r, column=4, value=row["Sales"])
        ws.cell(row=r, column=5, value=note)
        _style_data(ws, r, len(hdrs), bg, fg, num_cols={4})
        ws.cell(row=r, column=1).number_format = "DD-MMM-YYYY"
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")

    fr = 3 + len(df)
    ws.cell(row=fr, column=3, value="TOTAL")
    ws.cell(row=fr, column=4, value=df["Sales"].sum())
    for c in range(1, len(hdrs) + 1):
        cell = ws.cell(row=fr, column=c)
        cell.fill   = _fill(TOT_BG)
        cell.font   = _font(TOT_FG, bold=True)
        cell.border = MED
        cell.alignment = Alignment(horizontal="center" if c != 4 else "right")
    ws.cell(row=fr, column=4).number_format = "#,##0"
    ws.freeze_panes = "A3"


def _build_filtered_sheet(wb, tenants_data, day_type, year, month):
    is_we      = day_type == "Weekend"
    sheet_name = "All_Tenants_Weekends" if is_we else "All_Tenants_Weekdays"
    month_label = date(year, month, 1).strftime("%B %Y")

    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value     = f"All Tenants — {day_type} Days ({month_label})"
    t.fill      = _fill(ACC_BG if is_we else HDR_BG)
    t.font      = Font(color="FFFFFF", bold=True, size=14, name="Calibri")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    hdrs = ["Tenant", "Date", "Day Name", "Day Type", "Sales (IDR)", "Running Total"]
    for c, h in enumerate(hdrs, 1):
        ws.cell(row=2, column=c, value=h)
    _style_header(ws, 2, len(hdrs))
    _set_widths(ws, {1: 18, 2: 16, 3: 14, 4: 14, 5: 20, 6: 20})

    combined = []
    for name, df in tenants_data.items():
        filt = df[df["DayType"] == day_type].copy()
        filt.insert(0, "Tenant", name.replace("_", " "))
        combined.append(filt)

    if not combined:
        return

    all_df = pd.concat(combined, ignore_index=True)
    all_df.sort_values(["Date", "Tenant"], inplace=True)

    running = 0
    for i, (_, row) in enumerate(all_df.iterrows()):
        r       = 3 + i
        running += row["Sales"]
        bg      = WE_BG if is_we else (ALT_BG if i % 2 == 0 else "FFFFFF")
        fg      = WE_FG if is_we else "2C3E50"

        ws.cell(row=r, column=1, value=row["Tenant"])
        ws.cell(row=r, column=2, value=row["Date"])
        ws.cell(row=r, column=3, value=row["DayName"])
        ws.cell(row=r, column=4, value=row["DayType"])
        ws.cell(row=r, column=5, value=row["Sales"])
        ws.cell(row=r, column=6, value=running)
        _style_data(ws, r, len(hdrs), bg, fg, num_cols={5, 6})
        ws.cell(row=r, column=2).number_format = "DD-MMM-YYYY"

    ws.freeze_panes = "A3"


# ══════════════════════════════════════════════════════════════════
#  SECTION 5 — Chart Injection
# ══════════════════════════════════════════════════════════════════

def _inject_chart(wb, tenant_name: str, n_days: int) -> None:
    ws = wb[tenant_name[:31]]

    data_start = 3
    data_end   = data_start + n_days - 1

    sales_ref = Reference(ws, min_col=4, max_col=4,
                          min_row=data_start, max_row=data_end)
    date_ref  = Reference(ws, min_col=1, max_col=1,
                          min_row=data_start, max_row=data_end)

    chart = LineChart()
    chart.title          = f"{tenant_name.replace('_', ' ')} — Daily Sales"
    chart.style          = 10
    chart.y_axis.title   = "Sales (IDR)"
    chart.x_axis.title   = "Date"
    chart.y_axis.numFmt  = "#,##0"
    chart.width          = 24
    chart.height         = 14

    chart.add_data(sales_ref, titles_from_data=False)
    chart.set_categories(date_ref)

    series = chart.series[0]
    series.title  = SeriesLabel(v=tenant_name.replace("_", " "))
    series.smooth = True
    series.graphicalProperties.line.solidFill = "1A5276"
    series.graphicalProperties.line.width     = 22000
    series.marker.symbol = "circle"
    series.marker.size   = 5
    series.marker.graphicalProperties.fgColor              = "2874A6"
    series.marker.graphicalProperties.line.solidFill        = "2874A6"

    ws.add_chart(chart, f"{get_column_letter(7)}2")


# ══════════════════════════════════════════════════════════════════
#  SECTION 6 — Master Export
# ══════════════════════════════════════════════════════════════════

def export_report(
    tenants_data: dict,
    year: int,
    month: int,
    output_path: Path,
) -> Path:
    n_days = cal.monthrange(year, month)[1]

    wb = Workbook()
    wb.remove(wb.active)

    _build_summary(wb, tenants_data, year, month)

    for name, df in tenants_data.items():
        _build_tenant_sheet(wb, name, df, year, month)

    _build_filtered_sheet(wb, tenants_data, "Weekend", year, month)
    _build_filtered_sheet(wb, tenants_data, "Weekday", year, month)

    # Save once, then inject charts
    tmp = output_path.with_suffix(".tmp.xlsx")
    wb.save(tmp)

    wb2 = load_workbook(tmp)
    for name in tenants_data:
        if name[:31] in wb2.sheetnames:
            _inject_chart(wb2, name, n_days)
    wb2.save(output_path)
    tmp.unlink(missing_ok=True)

    return output_path


# ══════════════════════════════════════════════════════════════════
#  SECTION 7 — Dashboard Payload Builder
# ══════════════════════════════════════════════════════════════════

def build_dashboard(tenants_data: dict, year: int, month: int) -> dict:
    summary = []
    charts  = {}

    for name, df in tenants_data.items():
        total  = float(df["Sales"].sum())
        wd_s   = df[df["DayType"] == "Weekday"]["Sales"]
        we_s   = df[df["DayType"] == "Weekend"]["Sales"]
        wd_avg = float(wd_s.mean()) if len(wd_s) else 0.0
        we_avg = float(we_s.mean()) if len(we_s) else 0.0
        ratio  = round(we_avg / wd_avg, 2) if wd_avg else 0.0

        summary.append({
            "tenant"  : name.replace("_", " "),
            "total"   : total,
            "wd_avg"  : round(wd_avg, 0),
            "we_avg"  : round(we_avg, 0),
            "ratio"   : ratio,
            "wd_days" : int(len(wd_s)),
            "we_days" : int(len(we_s)),
        })

        charts[name.replace("_", " ")] = {
            "labels" : [str(d) for d in df["Date"]],
            "values" : [float(v) for v in df["Sales"]],
            "types"  : list(df["DayType"]),
        }

    return {"summary": summary, "charts": charts}


# ══════════════════════════════════════════════════════════════════
#  SECTION 8 — HTML Template
# ══════════════════════════════════════════════════════════════════

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Tenant Sales Aggregator</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    :root {
      --bg:      #0f172a; --surface: #1e293b; --surface2: #263248;
      --border:  rgba(255,255,255,0.08);
      --accent:  #3b82f6; --green:   #10b981; --danger: #ef4444;
      --text:    #e2e8f0; --muted:   #64748b;
    }
    html { scroll-behavior: smooth; }
    body { font-family: 'Segoe UI', system-ui, sans-serif;
           background: var(--bg); color: var(--text); min-height: 100vh; }

    /* ── topbar ── */
    .topbar {
      background: linear-gradient(90deg,#1e3a5f,#1a1a2e);
      padding: 1rem 2rem; display: flex; align-items: center;
      justify-content: space-between;
      border-bottom: 1px solid var(--border);
      position: sticky; top: 0; z-index: 100;
    }
    .logo { font-size: 1.2rem; font-weight: 700; color: #fff; }
    .logo span { color: var(--accent); }
    .tagline { color: var(--muted); font-size: 0.8rem; }

    /* ── layout ── */
    .main { max-width: 1100px; margin: 0 auto; padding: 2rem 1.5rem 4rem; }

    /* ── cards ── */
    .card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: 14px; padding: 1.5rem; margin-bottom: 1.5rem;
    }
    .card-title {
      font-size: 1rem; font-weight: 600; color: #93c5fd;
      margin-bottom: 1.25rem; display: flex; align-items: center; gap: .5rem;
    }

    /* ── controls ── */
    .ctrl-row { display: flex; gap: 1rem; flex-wrap: wrap; align-items: flex-end; }
    .form-group label { display: block; font-size: .8rem; color: var(--muted); margin-bottom: .3rem; }
    .form-group select, .form-group input[type=number] {
      background: var(--surface2); border: 1px solid var(--border);
      border-radius: 8px; color: var(--text);
      padding: .5rem .85rem; font-size: .95rem; outline: none;
    }
    .form-group select option { background: var(--surface); }

    /* ── drop zone ── */
    .dz {
      border: 2px dashed rgba(59,130,246,.35); border-radius: 12px;
      padding: 2.5rem 2rem; text-align: center; cursor: pointer;
      transition: all .2s;
    }
    .dz:hover, .dz.over {
      border-color: var(--accent); background: rgba(59,130,246,.05);
    }
    .dz .icon { font-size: 2.5rem; margin-bottom: .5rem; }
    .dz p { color: var(--muted); font-size: .9rem; margin-top: .25rem; }
    .dz strong { color: #93c5fd; }
    #fileInput { display: none; }

    /* ── chips ── */
    .chips { display: flex; flex-wrap: wrap; gap: .5rem; margin-top: .75rem; }
    .chip {
      display: inline-flex; align-items: center; gap: .4rem;
      background: var(--surface2); border: 1px solid var(--border);
      border-radius: 20px; padding: .3rem .75rem; font-size: .82rem;
    }
    .chip-x { cursor: pointer; color: var(--danger); border: none;
               background: none; font-size: .9rem; }

    /* ── buttons ── */
    .btn-row { display: flex; gap: .75rem; flex-wrap: wrap; margin-top: 1.25rem; }
    .btn {
      display: inline-flex; align-items: center; gap: .45rem;
      padding: .65rem 1.5rem; border-radius: 9px; font-size: .95rem;
      font-weight: 600; border: none; cursor: pointer; transition: all .2s;
    }
    .btn:disabled { opacity: .38; cursor: not-allowed; pointer-events: none; }
    .btn-p { background: var(--accent); color: #fff; }
    .btn-p:hover { background: #2563eb; transform: translateY(-1px); }
    .btn-g { background: var(--green); color: #fff; }
    .btn-g:hover { background: #059669; transform: translateY(-1px); }
    .btn-o { background: transparent; color: var(--text);
              border: 1px solid var(--border); }
    .btn-o:hover { background: var(--surface2); }

    /* ── alerts ── */
    .alert {
      display: flex; align-items: flex-start; gap: .75rem;
      padding: .9rem 1.1rem; border-radius: 10px; font-size: .9rem; margin-top: 1rem;
    }
    .alert.hidden  { display: none; }
    .alert-info    { background:rgba(59,130,246,.1);  border:1px solid rgba(59,130,246,.3);  color:#93c5fd; }
    .alert-success { background:rgba(16,185,129,.1);  border:1px solid rgba(16,185,129,.3);  color:#6ee7b7; }
    .alert-error   { background:rgba(239,68,68,.1);   border:1px solid rgba(239,68,68,.3);   color:#fca5a5; }
    .alert-warn    { background:rgba(251,191,36,.1);  border:1px solid rgba(251,191,36,.3);  color:#fde68a; }
    .spin {
      width:16px; height:16px; flex-shrink:0;
      border:2px solid transparent; border-top-color:currentColor;
      border-radius:50%; animation:spin .7s linear infinite;
    }
    @keyframes spin { to { transform:rotate(360deg); } }

    /* ── stat cards ── */
    #dash { display: none; }
    .stat-grid {
      display: grid; grid-template-columns: repeat(auto-fit,minmax(180px,1fr));
      gap: 1rem; margin-bottom: 1.5rem;
    }
    .stat-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: 12px; padding: 1.1rem 1.25rem;
    }
    .stat-card.green { border-color: rgba(16,185,129,.3); }
    .stat-card.blue  { border-color: rgba(59,130,246,.3); }
    .stat-card .lbl  { font-size:.78rem; color:var(--muted); margin-bottom:.4rem; }
    .stat-card .val  { font-size:1.5rem; font-weight:700; }
    .stat-card .sub  { font-size:.78rem; color:var(--muted); margin-top:.25rem; }

    /* ── table ── */
    .tbl-wrap { overflow-x:auto; border-radius:10px; border:1px solid var(--border); }
    table { width:100%; border-collapse:collapse; font-size:.88rem; }
    thead th {
      background:#1a3a5c; color:#93c5fd; padding:.75rem 1rem;
      text-align:left; font-weight:600; white-space:nowrap;
      border-bottom:1px solid var(--border);
    }
    thead th:not(:first-child) { text-align:right; }
    tbody tr { border-bottom:1px solid var(--border); transition:background .15s; }
    tbody tr:hover { background:var(--surface2); }
    tbody tr:last-child { background:rgba(59,130,246,.08); font-weight:700; border:none; }
    tbody td { padding:.65rem 1rem; }
    tbody td:not(:first-child) { text-align:right; }
    .badge-we {
      background:rgba(16,185,129,.12); border:1px solid rgba(16,185,129,.3);
      color:#6ee7b7; border-radius:4px; padding:.1rem .4rem; font-size:.75rem; font-weight:600;
    }
    .badge-wd {
      background:rgba(59,130,246,.12); border:1px solid rgba(59,130,246,.3);
      color:#93c5fd; border-radius:4px; padding:.1rem .4rem; font-size:.75rem; font-weight:600;
    }

    /* ── scan report ── */
    .scan-row { font-size:.85rem; }
    .scan-row code {
      background:rgba(255,255,255,.07); padding:.1rem .35rem;
      border-radius:4px; font-size:.8rem;
    }

    /* ── charts ── */
    .tabs { display:flex; gap:.4rem; flex-wrap:wrap; margin-bottom:1rem; }
    .tab {
      padding:.4rem 1rem; border-radius:20px; border:1px solid var(--border);
      background:transparent; color:var(--muted); font-size:.85rem; cursor:pointer; transition:all .15s;
    }
    .tab.active { background:var(--accent); border-color:var(--accent); color:#fff; }
    .pane { display:none; }
    .pane.active { display:block; }
    .chart-box { position:relative; height:320px; width:100%; }

    footer { text-align:center; padding:1.5rem; color:var(--muted);
             font-size:.8rem; border-top:1px solid var(--border); }
  </style>
</head>
<body>

<nav class="topbar">
  <div class="logo">📊 Tenant<span>Sales</span></div>
  <div class="tagline">Indonesian Calendar · Smart Scanner · OpenPyXL Export</div>
</nav>

<div class="main">

  <!-- Config -->
  <div class="card">
    <div class="card-title">⚙️ Configuration</div>
    <div class="ctrl-row">
      <div class="form-group">
        <label>Report Month</label>
        <select id="monthSel">
          <option value="1">January</option><option value="2">February</option>
          <option value="3">March</option><option value="4">April</option>
          <option value="5">May</option><option value="6">June</option>
          <option value="7">July</option><option value="8">August</option>
          <option value="9">September</option><option value="10">October</option>
          <option value="11">November</option><option value="12">December</option>
        </select>
      </div>
      <div class="form-group">
        <label>Year</label>
        <input id="yearIn" type="number" value="2025" min="2000" max="2100" style="width:90px"/>
      </div>
    </div>
  </div>

  <!-- Upload -->
  <div class="card">
    <div class="card-title">📁 Upload Tenant Files</div>
    <div class="dz" id="dz">
      <div class="icon">🗂️</div>
      <p><strong>Drag &amp; drop</strong> Excel files here</p>
      <p>or <strong style="text-decoration:underline;cursor:pointer"
          onclick="document.getElementById('fileInput').click()">click to browse</strong></p>
      <p style="font-size:.78rem;color:var(--muted);margin-top:.5rem">
        .xlsx or .xls · Headers can be anywhere in the file · Indonesian or English column names OK
      </p>
      <input type="file" id="fileInput" accept=".xlsx,.xls" multiple/>
    </div>
    <div class="chips" id="chips"></div>
    <div class="btn-row">
      <button class="btn btn-p" id="procBtn" disabled onclick="processFiles()">
        ⚙️ Process Files
      </button>
      <button class="btn btn-g" id="dlBtn" disabled onclick="downloadReport()">
        ⬇️ Export Excel Report
      </button>
      <button class="btn btn-o" onclick="resetAll()">↺ Reset</button>
    </div>
    <div class="alert hidden" id="alertBox">
      <span id="alertIcon"></span>
      <span id="alertMsg"></span>
    </div>
  </div>

  <!-- Dashboard -->
  <div id="dash">
    <div class="stat-grid" id="statGrid"></div>

    <!-- Scan report -->
    <div class="card" id="scanCard">
      <div class="card-title">🔍 File Scan Report</div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>File</th><th style="text-align:left">Status</th>
              <th style="text-align:left">Header Row</th>
              <th style="text-align:left">Date Column</th>
              <th style="text-align:left">Sales Column</th>
              <th style="text-align:left">Notes</th>
            </tr>
          </thead>
          <tbody id="scanBody"></tbody>
        </table>
      </div>
    </div>

    <!-- Summary table -->
    <div class="card">
      <div class="card-title">📋 Monthly Summary</div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Tenant</th><th>Total Sales (IDR)</th>
              <th>Weekday Days</th><th>Weekday Avg/Day</th>
              <th>Weekend Days</th><th>Weekend Avg/Day</th>
              <th>WE/WD Ratio</th>
            </tr>
          </thead>
          <tbody id="summaryBody"></tbody>
        </table>
      </div>
    </div>

    <!-- Charts -->
    <div class="card">
      <div class="card-title">📈 Daily Sales Charts</div>
      <div class="tabs" id="tabs"></div>
      <div id="panes"></div>
    </div>
  </div>

</div>

<footer>Tenant Sales Aggregator · Smart Column Scanner · Indonesian Holiday Logic</footer>

<script>
// ── state ──────────────────────────────────────────────────────────
let files = [], sessionId = null, charts = {};
document.getElementById("monthSel").value = new Date().getMonth() + 1;

// ── drop zone ──────────────────────────────────────────────────────
const dz = document.getElementById("dz");
const fi = document.getElementById("fileInput");
dz.addEventListener("dragover",  e => { e.preventDefault(); dz.classList.add("over"); });
dz.addEventListener("dragleave", ()=> dz.classList.remove("over"));
dz.addEventListener("drop", e => { e.preventDefault(); dz.classList.remove("over"); addFiles([...e.dataTransfer.files]); });
fi.addEventListener("change", () => { addFiles([...fi.files]); fi.value=""; });

function addFiles(newFiles) {
  const ok = [".xlsx", ".xls"];
  newFiles.forEach(f => {
    const ext = f.name.slice(f.name.lastIndexOf(".")).toLowerCase();
    if (!ok.includes(ext)) { showAlert(`"${f.name}" must be .xlsx or .xls`, "error"); return; }
    if (!files.find(x => x.name === f.name)) files.push(f);
  });
  renderChips();
}
function removeFile(i) { files.splice(i,1); renderChips(); }
function renderChips() {
  document.getElementById("chips").innerHTML = files.map((f,i) => `
    <div class="chip">
      📄 ${f.name}
      <span style="color:var(--muted);font-size:.75rem">${(f.size/1024).toFixed(0)}KB</span>
      <button class="chip-x" onclick="removeFile(${i})">✕</button>
    </div>`).join("");
  document.getElementById("procBtn").disabled = files.length === 0;
  document.getElementById("dlBtn").disabled   = true;
  sessionId = null;
  hideAlert();
}

// ── process ────────────────────────────────────────────────────────
async function processFiles() {
  if (!files.length) return;
  const month = document.getElementById("monthSel").value;
  const year  = document.getElementById("yearIn").value;

  showAlert("⏳ Scanning and processing files…", "info", true);
  document.getElementById("procBtn").disabled = true;
  document.getElementById("dash").style.display = "none";

  const form = new FormData();
  files.forEach(f => form.append("files", f));
  form.append("month", month);
  form.append("year",  year);

  try {
    const res  = await fetch("/api/process", { method:"POST", body:form });
    const data = await res.json();

    if (!res.ok || data.error) {
      showAlert(data.error || "Processing failed.", "error");
      renderScanReport(data.scan_reports);
      return;
    }

    sessionId = data.session_id;
    document.getElementById("dlBtn").disabled = false;
    showAlert(
      `✅ Done — ${data.tenant_count} tenant(s), ${data.total_rows} total rows processed.`,
      "success"
    );
    renderScanReport(data.scan_reports);
    renderDashboard(data.dashboard);

  } catch(err) {
    showAlert("Network error: " + err.message, "error");
  } finally {
    document.getElementById("procBtn").disabled = false;
  }
}

// ── download ───────────────────────────────────────────────────────
function downloadReport() {
  if (sessionId) window.location.href = `/api/download/${sessionId}`;
}

// ── reset ──────────────────────────────────────────────────────────
function resetAll() {
  files = []; sessionId = null;
  renderChips();
  document.getElementById("dash").style.display = "none";
  document.getElementById("dlBtn").disabled = true;
  Object.values(charts).forEach(c => c.destroy());
  charts = {};
  hideAlert();
}

// ── scan report ────────────────────────────────────────────────────
function renderScanReport(reports) {
  if (!reports || !reports.length) return;
  const body = document.getElementById("scanBody");
  body.innerHTML = reports.map(r => {
    const ok   = r.success;
    const stat = ok
      ? `<span style="color:#6ee7b7">✅ OK</span>`
      : `<span style="color:#fca5a5">❌ Failed</span>`;
    const warns = r.warnings && r.warnings.length
      ? `<br><span style="color:#fde68a;font-size:.78rem">⚠ ${r.warnings.join("<br>⚠ ")}</span>`
      : "";
    return `
      <tr class="scan-row">
        <td><strong>${r.file}</strong></td>
        <td>${stat}</td>
        <td>${ok ? `Row <code>${r.header_row}</code>` : `<span style="color:var(--muted)">—</span>`}</td>
        <td>${ok ? `<code>${r.date_col}</code>`  : `<span style="color:#fca5a5;font-size:.8rem">${r.error}</span>`}</td>
        <td>${ok ? `<code>${r.sales_col}</code>` : "—"}</td>
        <td>${warns || '<span style="color:var(--muted)">—</span>'}</td>
      </tr>`;
  }).join("");
  document.getElementById("dash").style.display = "block";
}

// ── dashboard ──────────────────────────────────────────────────────
function renderDashboard(data) {
  if (!data) return;
  const { summary, charts: chartData } = data;

  // stat cards
  const grandTotal = summary.reduce((s,r)=>s+r.total, 0);
  const top        = [...summary].sort((a,b)=>b.total-a.total)[0];
  const avgRatio   = summary.reduce((s,r)=>s+r.ratio,0) / summary.length;
  document.getElementById("statGrid").innerHTML = `
    <div class="stat-card blue">
      <div class="lbl">TENANTS LOADED</div>
      <div class="val">${summary.length}</div>
      <div class="sub">Files parsed</div>
    </div>
    <div class="stat-card">
      <div class="lbl">GRAND TOTAL SALES</div>
      <div class="val">${fmtShort(grandTotal)}</div>
      <div class="sub">IDR all tenants</div>
    </div>
    <div class="stat-card green">
      <div class="lbl">AVG WEEKEND RATIO</div>
      <div class="val">${avgRatio.toFixed(2)}×</div>
      <div class="sub">vs weekday average</div>
    </div>
    <div class="stat-card">
      <div class="lbl">TOP TENANT</div>
      <div class="val" style="font-size:1.1rem">${top?.tenant||"—"}</div>
      <div class="sub">${fmtShort(top?.total||0)} IDR</div>
    </div>`;

  // summary table
  let gt = 0;
  document.getElementById("summaryBody").innerHTML =
    summary.map(r => {
      gt += r.total;
      return `<tr>
        <td><strong>${r.tenant}</strong></td>
        <td>${fmtIDR(r.total)}</td>
        <td><span class="badge-wd">${r.wd_days}d</span></td>
        <td>${fmtIDR(r.wd_avg)}</td>
        <td><span class="badge-we">${r.we_days}d</span></td>
        <td>${fmtIDR(r.we_avg)}</td>
        <td>${r.ratio.toFixed(2)}×</td>
      </tr>`;
    }).join("") +
    `<tr>
      <td>🏆 Grand Total</td><td>${fmtIDR(gt)}</td>
      <td colspan="5" style="text-align:center;color:var(--muted)">— all tenants combined —</td>
    </tr>`;

  // charts
  const tabsEl  = document.getElementById("tabs");
  const panesEl = document.getElementById("panes");
  tabsEl.innerHTML = ""; panesEl.innerHTML = "";
  Object.values(charts).forEach(c => c.destroy());
  charts = {};

  Object.entries(chartData).forEach(([tenant, cd], idx) => {
    const id  = "c_" + tenant.replace(/\W/g,"_");

    const tab = document.createElement("button");
    tab.className   = "tab" + (idx===0?" active":"");
    tab.textContent = tenant;
    tab.onclick     = () => switchTab(id, tab);
    tabsEl.appendChild(tab);

    const pane = document.createElement("div");
    pane.className = "pane" + (idx===0?" active":"");
    pane.id = id;
    pane.innerHTML = `<div class="chart-box"><canvas id="cv_${id}"></canvas></div>`;
    panesEl.appendChild(pane);

    const labels   = cd.labels.map(d => {
      const dt = new Date(d);
      return dt.toLocaleDateString("en-GB",{day:"2-digit",month:"short"});
    });
    const ptColors = cd.types.map(t => t==="Weekend"?"#10b981":"#3b82f6");

    const ctx = document.getElementById(`cv_${id}`).getContext("2d");
    charts[id] = new Chart(ctx, {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: `${tenant} Daily Sales (IDR)`,
          data: cd.values,
          borderColor: "#3b82f6",
          borderWidth: 2,
          pointBackgroundColor: ptColors,
          pointBorderColor: ptColors,
          pointRadius: 4,
          pointHoverRadius: 6,
          fill: true,
          backgroundColor: ctx2 => {
            const chart = ctx2.chart;
            if (!chart.chartArea) return "transparent";
            const g = ctx2.chart.ctx.createLinearGradient(
              0, chart.chartArea.top, 0, chart.chartArea.bottom
            );
            g.addColorStop(0,"rgba(59,130,246,.25)");
            g.addColorStop(1,"rgba(59,130,246,0)");
            return g;
          },
          tension: 0.35,
        }],
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        interaction:{ mode:"index", intersect:false },
        plugins:{
          legend:{ labels:{ color:"#94a3b8", font:{size:12} } },
          tooltip:{
            backgroundColor:"#1e293b", borderColor:"#334155", borderWidth:1,
            titleColor:"#93c5fd", bodyColor:"#e2e8f0",
            callbacks:{
              label: ctx => ` ${fmtIDR(ctx.parsed.y)} IDR`,
              afterLabel: ctx => cd.types[ctx.dataIndex]==="Weekend"
                ? "  🟢 Weekend / Holiday" : "  🔵 Weekday",
            }
          },
        },
        scales:{
          x:{ ticks:{color:"#64748b",maxRotation:45,font:{size:10}},
              grid:{color:"rgba(255,255,255,.04)"} },
          y:{ ticks:{color:"#64748b", callback:v=>fmtShort(v)},
              grid:{color:"rgba(255,255,255,.06)"} },
        },
      },
    });
  });
}

function switchTab(id, btn) {
  document.querySelectorAll(".pane").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(b=>b.classList.remove("active"));
  document.getElementById(id).classList.add("active");
  btn.classList.add("active");
}

// ── alerts ─────────────────────────────────────────────────────────
function showAlert(msg, type="info", spinner=false) {
  const box  = document.getElementById("alertBox");
  const icon = document.getElementById("alertIcon");
  const text = document.getElementById("alertMsg");
  box.className = `alert alert-${type}`;
  text.innerHTML = msg;
  icon.innerHTML = spinner ? '<div class="spin"></div>'
    : type==="error" ? "⚠️" : type==="success" ? "✅" : "ℹ️";
}
function hideAlert() {
  document.getElementById("alertBox").className = "alert hidden";
}

// ── formatters ─────────────────────────────────────────────────────
function fmtIDR(v) {
  return Number(v).toLocaleString("id-ID");
}
function fmtShort(v) {
  if (v>=1e9) return (v/1e9).toFixed(2)+"B";
  if (v>=1e6) return (v/1e6).toFixed(2)+"M";
  if (v>=1e3) return (v/1e3).toFixed(0)+"K";
  return v;
}
</script>
</body>
</html>
"""


# ══════════════════════════════════════════════════════════════════
#  SECTION 9 — Flask Routes
# ══════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/process", methods=["POST"])
def api_process():
    if "files" not in request.files:
        return jsonify({"error": "No files uploaded."}), 400

    try:
        month = int(request.form.get("month", 1))
        year  = int(request.form.get("year", datetime.now().year))
        if not (1 <= month <= 12):
            raise ValueError
        if not (2000 <= year <= 2100):
            raise ValueError
    except ValueError:
        return jsonify({"error": "Month must be 1–12, year must be 2000–2100."}), 400

    uploaded = request.files.getlist("files")
    if not uploaded or all(f.filename == "" for f in uploaded):
        return jsonify({"error": "No files received."}), 400

    session_id  = str(uuid.uuid4())
    session_dir = UPLOAD_FOLDER / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for f in uploaded:
        if not f.filename:
            continue
        if not allowed_file(f.filename):
            return jsonify({
                "error": f'"{f.filename}" is not a valid Excel file (.xlsx / .xls).'
            }), 400
        dest = session_dir / secure_filename(f.filename)
        f.save(dest)
        saved.append(dest)

    if not saved:
        return jsonify({"error": "No valid files were saved."}), 400

    try:
        tenants_data, scan_reports = parse_all_tenants(saved, year, month)

        if not tenants_data:
            details = "; ".join(
                f"{r['file']}: {r['error']}" for r in scan_reports
            )
            return jsonify({
                "error"        : f"No valid data parsed. {details}",
                "scan_reports" : scan_reports,
            }), 422

        out_path = session_dir / f"Tenant_Report_{year}_{month:02d}.xlsx"
        export_report(tenants_data, year, month, out_path)

        dashboard = build_dashboard(tenants_data, year, month)
        SESSION_STORE[session_id] = {"path": out_path}

        total_rows = sum(len(df) for df in tenants_data.values())

        return jsonify({
            "session_id"   : session_id,
            "tenant_count" : len(tenants_data),
            "total_rows"   : total_rows,
            "dashboard"    : dashboard,
            "scan_reports" : scan_reports,
        })

    except Exception as exc:
        app.logger.exception("Processing error")
        return jsonify({"error": f"Processing error: {str(exc)}"}), 500


@app.route("/api/download/<session_id>")
def api_download(session_id: str):
    entry = SESSION_STORE.get(session_id)
    if not entry:
        return jsonify({"error": "Report not found. Please process your files again."}), 404
    path: Path = entry["path"]
    if not path.exists():
        return jsonify({"error": "Report file missing on server."}), 500
    return send_file(
        path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=path.name,
    )


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  🚀  Starting on http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
