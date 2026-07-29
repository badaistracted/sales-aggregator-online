# app.py
import os
import re
import uuid
import calendar as cal
from datetime import datetime, date
import pandas as pd
import pdfplumber
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = Path("temp_uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)

# ══════════════════════════════════════════════════════════════
#  MONTH REFERENCE DATA
# ══════════════════════════════════════════════════════════════

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

DATE_ALIASES = [
    "date", "tanggal", "tgl", "hari", "periode", "period",
    "transaction date", "trans date", "waktu", "datetime",
    "tgl transaksi", "tanggal transaksi",
]

SALES_ALIASES = [
    "sales", "penjualan", "total", "amount", "revenue",
    "total sales", "total penjualan", "jumlah", "nilai",
    "net sales", "gross sales", "omzet", "omset",
    "total amount", "sales amount", "income", "pendapatan",
]


# ══════════════════════════════════════════════════════════════
#  MONTH DETECTION (for validation — from Phase 1.6)
# ══════════════════════════════════════════════════════════════

def detect_months_in_text(text):
    text = str(text).lower().strip()
    found = set()
    if not text or text in ("nan", "none"):
        return found
    clean = text.replace(",", "").replace(".", "").strip()
    if clean.isdigit():
        return found

    for match in re.finditer(r"([a-z]{3,})[\s\-_/\.]+(\d{2,4})", text):
        name, yr = match.group(1), int(match.group(2))
        if yr < 100: yr += 2000
        mn = MONTH_NAMES.get(name)
        if mn and 2020 <= yr <= 2035:
            found.add((yr, mn))

    for match in re.finditer(r"(\d{4})[\s\-_/\.]+(\d{1,2})(?!\d)", text):
        yr, mn = int(match.group(1)), int(match.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            found.add((yr, mn))

    for match in re.finditer(r"(\d{4})[\s\-_/\.]+([a-z]{3,})", text):
        yr = int(match.group(1))
        mn = MONTH_NAMES.get(match.group(2))
        if mn and 2020 <= yr <= 2035:
            found.add((yr, mn))

    for match in re.finditer(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", text):
        a, b, c = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 2020 <= c <= 2035:
            if 1 <= b <= 12: found.add((c, b))
            if 1 <= a <= 12 and a != b: found.add((c, a))

    for match in re.finditer(r"(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})", text):
        yr, mn = int(match.group(1)), int(match.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            found.add((yr, mn))

    return found


def detect_months_in_excel(rows):
    found = set()
    for row in rows:
        for cell in row:
            s = str(cell).strip()
            if s and s not in ("", "nan", "None"):
                found.update(detect_months_in_text(s))
    return found


def detect_months_in_lines(lines):
    found = set()
    for line in lines:
        found.update(detect_months_in_text(line))
    return found


def validate_month(detected, target_year, target_month):
    target = (target_year, target_month)
    target_label = cal.month_name[target_month] + " " + str(target_year)

    if not detected:
        return {
            "status": "warning", "icon": "⚠️",
            "message": "No month/year detected. Cannot verify.",
            "match": False, "detected": [], "target": target_label,
        }

    detected_labels = sorted([cal.month_name[m] + " " + str(y) for y, m in detected])

    if target in detected:
        if len(detected) == 1:
            return {
                "status": "ok", "icon": "✅",
                "message": "Matches: " + target_label,
                "match": True, "detected": detected_labels, "target": target_label,
            }
        others = [l for l in detected_labels if l != target_label]
        return {
            "status": "ok_multi", "icon": "⚠️",
            "message": "Contains " + target_label + " but also has: " + ", ".join(others),
            "match": True, "detected": detected_labels, "target": target_label,
        }

    return {
        "status": "mismatch", "icon": "⚠️",
        "message": "WRONG MONTH — Expected " + target_label + " but found: " + ", ".join(detected_labels),
        "match": False, "detected": detected_labels, "target": target_label,
    }


# ══════════════════════════════════════════════════════════════
#  FILE READERS (Phase 1)
# ══════════════════════════════════════════════════════════════

def read_excel(path):
    try:
        engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine).fillna("")
        rows = []
        for row in df.values.tolist():
            rows.append([str(c) if str(c) != "" else "" for c in row])
        return {"type": "table", "rows": rows, "cols": len(rows[0]) if rows else 0}
    except Exception as e:
        return {"type": "error", "message": "Excel Error: " + str(e)}


def smart_ocr_to_table(lines):
    """Convert OCR lines to preview table. Only merges dates together."""
    table_rows = []
    date_pattern = re.compile(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Januari|Februari|Maret|Mei|Juni|Juli|Agustus|"
        r"Oktober|Desember)\s+(\d{4})",
        re.IGNORECASE
    )

    for line in lines:
        clean = str(line).strip()
        if not clean:
            continue
        merged = clean
        for day, month, year in date_pattern.findall(merged):
            original = day + " " + month + " " + year
            placeholder = day + "_" + month + "_" + year
            merged = merged.replace(original, placeholder, 1)

        parts = re.split(r"\s+", merged)
        restored = []
        for part in parts:
            if "_" in part and date_pattern.match(part.replace("_", " ")):
                restored.append(part.replace("_", " "))
            else:
                restored.append(part)
        table_rows.append(restored)

    if not table_rows:
        return None, 0
    max_cols = max(len(row) for row in table_rows)
    for row in table_rows:
        while len(row) < max_cols:
            row.append("")
    return table_rows, max_cols


def read_pdf_ocr(path):
    try:
        from pdf2image import convert_from_path
        import pytesseract

        images = convert_from_path(path, dpi=300)
        text_lines = []
        for img in images:
            text = pytesseract.image_to_string(img, lang="eng")
            if text:
                text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])

        if text_lines:
            rows, cols = smart_ocr_to_table(text_lines)
            if rows:
                return {
                    "type": "table", "rows": rows, "cols": cols,
                    "is_ocr": True, "raw_lines": text_lines,
                }
            return {"type": "text", "lines": text_lines}

        return {"type": "error", "message": "PDF OCR: No text detected."}
    except Exception as e:
        return {"type": "error", "message": "PDF OCR Error: " + str(e)}


def read_pdf(path):
    try:
        tables_found = []
        text_lines = []

        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        for row in table:
                            cleaned = [str(c).strip() if c else "" for c in row]
                            if any(cleaned):
                                tables_found.append(cleaned)
                text = page.extract_text()
                if text:
                    text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])

        if tables_found:
            max_cols = max(len(r) for r in tables_found)
            for r in tables_found:
                while len(r) < max_cols:
                    r.append("")
            return {
                "type": "table", "rows": tables_found, "cols": max_cols,
                "lines": text_lines,
            }

        if text_lines and len(text_lines) > 5:
            rows, cols = smart_ocr_to_table(text_lines)
            if rows:
                return {"type": "table", "rows": rows, "cols": cols, "lines": text_lines}
            return {"type": "text", "lines": text_lines}

        return read_pdf_ocr(path)

    except Exception:
        return read_pdf_ocr(path)


# ══════════════════════════════════════════════════════════════
#  PHASE 2 — PARSERS
# ══════════════════════════════════════════════════════════════

def parse_number(raw):
    """Parse Indonesian-style numbers: '330.685.175', '128,939,034', '5000000'."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--", "n/a"):
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


def parse_month_year_cell(text):
    """Parse a single cell as month-year header. Returns (year, month) or None."""
    if text is None:
        return None
    s = str(text).strip().lower()
    if not s or s in ("nan", "none"):
        return None

    m = re.fullmatch(r"([a-z]{3,})[\s\-_/\.]+(\d{2,4})", s)
    if m:
        mn = MONTH_NAMES.get(m.group(1))
        yr = int(m.group(2))
        if yr < 100: yr += 2000
        if mn and 2020 <= yr <= 2035:
            return (yr, mn)

    m = re.fullmatch(r"(\d{4})[\s\-_/\.]+(\d{1,2})", s)
    if m:
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            return (yr, mn)

    m = re.fullmatch(r"(\d{4})[\s\-_/\.]+([a-z]{3,})", s)
    if m:
        yr = int(m.group(1))
        mn = MONTH_NAMES.get(m.group(2))
        if mn and 2020 <= yr <= 2035:
            return (yr, mn)

    # Datetime strings from pandas: "2026-03-01 00:00:00"
    m = re.match(r"^(\d{4})-(\d{2})-\d{2}", s)
    if m:
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            return (yr, mn)

    return None


def _match_alias(cell, aliases):
    c = re.sub(r"\s+", " ", str(cell).strip().lower())
    if not c or c in ("nan", "none"):
        return False
    for a in aliases:
        if c == a or c.startswith(a) or a in c:
            return True
    return False

def _norm_header(x):
    return re.sub(r"\s+", " ", str(x).strip().lower())


def _score_date_header(cell):
    """
    Strict date-header scoring.
    IMPORTANT: does NOT treat 'periode' as a date column header.
    """
    c = _norm_header(cell)

    if c in ("tgl", "tanggal", "date"):
        return 100
    if c in ("transaction date", "trans date", "tanggal transaksi", "tgl transaksi"):
        return 95
    if c.startswith("tanggal ") or c.startswith("tgl "):
        return 90
    if c in ("datetime", "waktu"):
        return 60

    return 0


def _score_sales_header(cell):
    """
    Prioritise the best sales column, not the first one found.
    Example:
      TOTAL PENJUALAN  >  NETT PENJUALAN  >  PENJUALAN
    """
    c = _norm_header(cell)

    # strongest matches
    if c in ("total penjualan", "total sales", "gross sales", "sales total", "sales amount", "total amount"):
        return 100

    # strong partial matches
    if "total penjualan" in c or "total sales" in c:
        return 98

    # next-best choices
    if c in ("nett penjualan", "net penjualan", "net sales", "nett sales", "net revenue", "gross revenue"):
        return 90
    if "nett penjualan" in c or "net sales" in c or "net revenue" in c:
        return 88

    # still useful
    if c in ("pendapatan", "revenue", "omzet", "omset"):
        return 80
    if "pendapatan" in c or "revenue" in c or "omzet" in c or "omset" in c:
        return 78

    # weakest generic matches
    if c in ("penjualan", "sales", "amount", "nilai", "jumlah"):
        return 60
    if "penjualan" in c or "sales" in c:
        return 55

    return 0


def _finance_header_bonus(row):
    """
    Bonus for rows that look like real financial header rows.
    """
    bonus = 0
    for cell in row:
        c = _norm_header(cell)
        if any(tok in c for tok in (
            "penjualan", "sales", "total", "nett", "net",
            "revenue", "pendapatan", "fnb", "phi", "tax"
        )):
            bonus += 1
    return min(bonus, 6) * 4


def _candidate_data_score(rows, header_idx, date_c, sales_c):
    """
    Validate candidate header row using the rows below it.
    We want:
    - date column to parse as dates
    - sales column to parse as numbers
    """
    checked = 0
    good_dates = 0
    good_nums = 0

    for r in rows[header_idx + 1: header_idx + 13]:
        dv = r[date_c] if date_c < len(r) else ""
        sv = r[sales_c] if sales_c < len(r) else ""

        if str(dv).strip() == "" and str(sv).strip() == "":
            continue

        checked += 1

        dt = pd.to_datetime(str(dv).strip(), dayfirst=True, errors="coerce")
        if not pd.isna(dt):
            good_dates += 1

        if parse_number(sv) is not None:
            good_nums += 1

    if checked == 0:
        return 0

    return (good_dates / checked) * 40 + (good_nums / checked) * 40


# ─── Parser 1: Excel Pivot (tenants as rows, months as columns) ──

def try_excel_pivot(rows):
    best = None
    for idx, row in enumerate(rows[:50]):
        month_cols = {}
        for c, cell in enumerate(row):
            my = parse_month_year_cell(cell)
            if my:
                month_cols[c] = my
        if len(month_cols) >= 2:
            best = (idx, month_cols)
            break

    if not best:
        return None

    header_idx, month_cols = best
    ncols = max(len(r) for r in rows) if rows else 0

    label_col = None
    for c in range(ncols):
        if c not in month_cols:
            label_col = c
            break
    if label_col is None:
        return None

    skip_kw = ("total", "grand", "jumlah", "subtotal", "sub total",
               "sum", "rata", "average", "nan", "none")
    tenants = {}

    for r in rows[header_idx + 1:]:
        if label_col >= len(r):
            continue
        name = str(r[label_col]).strip()
        if not name or name.lower() in ("nan", "none"):
            continue
        if any(k in name.lower() for k in skip_kw):
            continue
        if not re.search(r"[a-zA-Z]", name):
            continue

        monthly = {}
        for c, (yr, mn) in month_cols.items():
            if c < len(r):
                v = parse_number(r[c])
                if v is not None:
                    key = f"{yr}-{mn:02d}"
                    monthly[key] = monthly.get(key, 0) + v
        if monthly:
            tenants[name] = {"monthly": monthly, "daily": []}

    if not tenants:
        return None

    months_found = sorted({k for t in tenants.values() for k in t["monthly"]})
    return {
        "success": True,
        "format": "excel_pivot",
        "tenants": tenants,
        "message": f"Pivot: {len(tenants)} tenant(s) × {len(months_found)} month(s). "
                   f"Header at row {header_idx + 1} (title rows above ignored).",
    }


# ─── Parser 2: Excel Columnar (Date | Sales columns) ──────────

def parse_date_cell(raw):
    """
    Parse dates safely.
    - If it looks like ISO format (YYYY-MM-DD), do NOT use dayfirst=True
    - Only use dayfirst=True for ambiguous human-style dates like 01/05/2026
    """
    if raw is None:
        return None

    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "-", "--"):
        return None

    iso_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d",
    ]
    for fmt in iso_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    if re.match(r"^\d{4}[-/]", s):
        dt = pd.to_datetime(s, errors="coerce", yearfirst=True)
        if not pd.isna(dt):
            return dt.date()

    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if not pd.isna(dt):
        return dt.date()

    return None

def try_excel_columnar(rows):
    """
    Detect a classic daily table with a true header row somewhere below title/logo rows.
    Example:
      TGL | PENJUALAN | FNB | CHOO-CHOO TRAIN | TOTAL PENJUALAN | ...
    """
    best = None

    for idx, row in enumerate(rows[:50]):
        date_candidates = []
        sales_candidates = []

        for c, cell in enumerate(row):
            ds = _score_date_header(cell)
            ss = _score_sales_header(cell)

            if ds > 0:
                date_candidates.append((ds, c, str(cell)))
            if ss > 0:
                sales_candidates.append((ss, c, str(cell)))

        if not date_candidates or not sales_candidates:
            continue

        best_date_score, date_c, date_label = max(date_candidates, key=lambda x: x[0])
        best_sales_score, sales_c, sales_label = max(sales_candidates, key=lambda x: x[0])

        score = (
            best_date_score +
            best_sales_score +
            _finance_header_bonus(row) +
            _candidate_data_score(rows, idx, date_c, sales_c)
        )

        if best is None or score > best["score"]:
            best = {
                "score": score,
                "header_idx": idx,
                "date_c": date_c,
                "sales_c": sales_c,
                "date_label": date_label,
                "sales_label": sales_label,
            }

    if best is None:
        return None

    header_idx = best["header_idx"]
    date_c     = best["date_c"]
    sales_c    = best["sales_c"]

    daily = []
    for r in rows[header_idx + 1:]:
        if date_c >= len(r) or sales_c >= len(r):
            continue

        raw_date = str(r[date_c]).strip()
        raw_sales = r[sales_c]

        parsed_date = parse_date_cell(raw_date)
if parsed_date is None:
    continue

daily.append({
    "date": str(parsed_date),
    "sales": s,
})
    if not daily:
        return None

    monthly = {}
    for d in daily:
        key = d["date"][:7]
        monthly[key] = monthly.get(key, 0) + d["sales"]

    return {
        "success": True,
        "format": "excel_columnar",
        "tenants": {
            "__FROM_FILENAME__": {
                "monthly": monthly,
                "daily": daily,
            }
        },
        "message": (
            f"Columnar: {len(daily)} daily rows. "
            f"Header row {header_idx + 1}. "
            f"Date='{best['date_label']}' | Sales='{best['sales_label']}'"
        ),
    }


# ─── Parser 3: PDF Daily Lines (OCR / digital text) ─────────

_MONTH_KEYS = sorted(MONTH_NAMES.keys(), key=len, reverse=True)
_MONTH_RX = "|".join(re.escape(k) for k in _MONTH_KEYS)

DAILY_RX = re.compile(
    r"^\s*\d{1,3}\s*[\.,]?\s+"           # row index: "1." or "11,"
    r"(\d{1,2})\s+"                       # day: 01
    r"(" + _MONTH_RX + r")\s+"            # month name: March
    r"(\d{4})"                            # year: 2026
    r"\s*,?\s*"
    r"([A-Za-z]+)?\s*"                    # optional day name: Sunday
    r"(.*)$",                             # rest: numbers + OCR noise
    re.IGNORECASE,
)

NUM_RX = re.compile(r"\d{1,3}(?:[.,]\d{3,})+|\d{5,}")


def try_pdf_daily(lines, filename=""):
    daily = []
    header_lines = []

    for line in lines:
        clean = str(line).strip()
        if not clean:
            continue
        m = DAILY_RX.match(clean)
        if m:
            day = int(m.group(1))
            mn = MONTH_NAMES.get(m.group(2).lower())
            yr = int(m.group(3))
            rest = m.group(5)

            if not mn or not (2020 <= yr <= 2035):
                continue
            try:
                d = date(yr, mn, day)
            except ValueError:
                continue

            nums = NUM_RX.findall(rest)
            vals = [parse_number(n) for n in nums]
            vals = [v for v in vals if v is not None]

            if vals:
                # Leftmost number = main "Total Penjualan" column
                daily.append({"date": str(d), "sales": vals[0]})
        else:
            header_lines.append(clean)

    if len(daily) < 3:
        return None

    # Tenant name: first line that isn't report metadata
    skip_kw = ("laporan", "report", "periode", "bulan", "page", "halaman",
               "untuk", "per mu", "per tanggal", "printed", "dicetak")
    tenant = None
    for hl in header_lines[:15]:
        low = hl.lower()
        if len(hl) < 3:
            continue
        if any(k in low for k in skip_kw):
            continue
        tenant = hl
        break

    if not tenant:
        tenant = Path(filename).stem if filename else "Unknown Tenant"

    monthly = {}
    for d in daily:
        key = d["date"][:7]
        monthly[key] = monthly.get(key, 0) + d["sales"]

    return {
        "success": True,
        "format": "pdf_daily",
        "tenants": {tenant: {"monthly": monthly, "daily": daily}},
        "message": f"Daily report: {len(daily)} days extracted for '{tenant}'. "
                   f"Sales = leftmost number column.",
    }


# ─── Master Parse Dispatcher ─────────────────────────────────

def parse_report(data, ext, filename):
    rows = data.get("rows") if data.get("type") == "table" else None
    lines = data.get("lines") or data.get("raw_lines")

    if rows:
        p = try_excel_pivot(rows)
        if p:
            return p
        p = try_excel_columnar(rows)
        if p:
            # Resolve filename placeholder
            stem = Path(filename).stem
            if "__FROM_FILENAME__" in p["tenants"]:
                p["tenants"][stem] = p["tenants"].pop("__FROM_FILENAME__")
            return p

    if lines:
        p = try_pdf_daily(lines, filename)
        if p:
            return p

    if rows:
        joined = [" ".join(str(c) for c in r if str(c).strip()) for r in rows]
        p = try_pdf_daily(joined, filename)
        if p:
            return p

    return {
        "success": False,
        "format": None,
        "tenants": {},
        "message": "Could not detect report structure (no pivot headers, "
                   "no Date/Sales columns, no daily rows).",
    }


# ══════════════════════════════════════════════════════════════
#  HTML UI
# ══════════════════════════════════════════════════════════════

HTML_UI = r"""
<!DOCTYPE html>
<html>
<head>
    <title>Tenant Report Reader</title>
    <style>
        body {
            font-family: sans-serif;
            background: #0f172a;
            color: #e2e8f0;
            padding: 40px;
            margin: 0;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { margin-bottom: 5px; }
        .subtitle { color: #94a3b8; margin-bottom: 25px; }

        .config-card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .config-title {
            font-size: 1rem;
            font-weight: 600;
            color: #93c5fd;
            margin-bottom: 15px;
        }
        .config-row { display: flex; gap: 15px; align-items: flex-end; flex-wrap: wrap; }
        .config-group label { display: block; font-size: 0.8rem; color: #94a3b8; margin-bottom: 5px; }
        .config-group select, .config-group input[type="number"] {
            background: #263248; border: 1px solid #475569; border-radius: 8px;
            color: #e2e8f0; padding: 8px 12px; font-size: 0.95rem; outline: none;
        }
        .config-group select option { background: #1e293b; }
        .month-preview { font-size: 0.85rem; color: #94a3b8; padding: 8px 0; }
        .month-preview b { color: #3b82f6; }

        .drop-zone {
            border: 3px dashed #3b82f6; border-radius: 20px; padding: 50px;
            text-align: center; cursor: pointer;
            background: rgba(59, 130, 246, 0.05); transition: 0.3s;
        }
        .drop-zone.hover {
            background: rgba(59, 130, 246, 0.15);
            border-color: #60a5fa; transform: scale(1.01);
        }
        .drop-zone .icon { font-size: 3rem; margin-bottom: 10px; }
        .drop-zone p { color: #94a3b8; margin: 5px 0; }
        .drop-zone b { color: #93c5fd; }

        .browse-row { display: flex; gap: 10px; justify-content: center; margin-top: 15px; flex-wrap: wrap; }
        .browse-btn {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 10px 20px; border-radius: 10px; font-size: 0.9rem;
            font-weight: 600; border: none; cursor: pointer; transition: all 0.2s;
        }
        .browse-btn:hover { transform: translateY(-1px); }
        .btn-files { background: #3b82f6; color: #fff; }
        .btn-files:hover { background: #2563eb; }
        .btn-folder { background: #10b981; color: #fff; }
        .btn-folder:hover { background: #059669; }
        .btn-clear { background: transparent; color: #94a3b8; border: 1px solid #475569; }
        .btn-clear:hover { background: #263248; }
        .hidden-input { display: none; }

        .loader {
            display: none; text-align: center; color: #3b82f6;
            font-weight: bold; margin-top: 15px; padding: 15px;
        }
        .spinner {
            display: inline-block; width: 14px; height: 14px;
            border: 2px solid #3b82f6; border-top: 2px solid transparent;
            border-radius: 50%; animation: spin 0.6s linear infinite;
            margin-right: 8px; vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .summary { margin-top: 20px; display: none; }
        .summary-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; }
        .summary-card {
            background: #1e293b; border: 1px solid #334155;
            padding: 15px; border-radius: 10px; text-align: center;
        }
        .summary-card .num { font-size: 1.5rem; font-weight: bold; }
        .summary-card .label { font-size: 0.75rem; color: #94a3b8; margin-top: 4px; }
        .c-blue { color: #3b82f6; } .c-green { color: #10b981; }
        .c-red { color: #ef4444; } .c-yellow { color: #f59e0b; }
        .c-purple { color: #a78bfa; }

        /* Master Report */
        .master { margin-top: 20px; display: none; }
        .master-warn {
            background: rgba(245,158,11,0.08);
            border: 1px solid rgba(245,158,11,0.25);
            color: #fde68a; padding: 10px 14px; border-radius: 8px;
            font-size: 0.85em; margin-bottom: 12px; display: none;
        }
        .master-table-wrap { overflow-x: auto; border-radius: 8px; border: 1px solid #334155; }
        .master-table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
        .master-table th {
            background: #1a3a5c; color: #93c5fd; padding: 8px 12px;
            text-align: right; font-weight: 600; white-space: nowrap;
            border-bottom: 2px solid #2563eb;
        }
        .master-table th:first-child, .master-table td:first-child { text-align: left; }
        .master-table td {
            padding: 7px 12px; border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #cbd5e1; text-align: right; white-space: nowrap;
        }
        .master-table tr:hover td { background: rgba(59,130,246,0.06); }
        .master-table .t-col { background: rgba(59,130,246,0.12); }
        .master-table .t-head { background: rgba(37,99,235,0.35); color: #fff; }
        .master-table .total-row td {
            background: rgba(59,130,246,0.1); font-weight: 700;
            border-top: 2px solid #3b82f6;
        }
        .master-table .src { font-size: 0.75em; color: #64748b; }
        .master-table .no-data { color: #334155; }

        /* File Cards */
        .file-list { margin-top: 20px; }
        .file-card {
            background: #1e293b; border: 1px solid #334155;
            border-radius: 12px; margin-bottom: 12px; overflow: hidden;
        }
        .file-card.mismatch { border-color: #ef4444; }
        .file-card.matched { border-color: #10b981; }
        .file-card.warning { border-color: #f59e0b; }
        .file-header {
            display: flex; justify-content: space-between; align-items: center;
            padding: 12px 15px; cursor: pointer; transition: background 0.2s;
        }
        .file-header:hover { background: #263248; }
        .file-info { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }

        .badge {
            font-size: 0.7em; padding: 3px 8px; border-radius: 4px;
            font-weight: bold; letter-spacing: 0.5px;
        }
        .b-xlsx { background: #2563eb; color: #fff; }
        .b-xls  { background: #7c3aed; color: #fff; }
        .b-pdf  { background: #f59e0b; color: #000; }
        .b-ok   { background: rgba(16,185,129,0.2); color: #10b981; border: 1px solid #10b981; }
        .b-fail { background: rgba(239,68,68,0.2); color: #ef4444; border: 1px solid #ef4444; }
        .b-parse { background: rgba(167,139,250,0.2); color: #a78bfa; border: 1px solid #a78bfa; }

        .month-badge { font-size: 0.72em; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
        .mb-match { background: rgba(16,185,129,0.15); color: #10b981; border: 1px solid rgba(16,185,129,0.3); }
        .mb-mismatch { background: rgba(239,68,68,0.15); color: #ef4444; border: 1px solid rgba(239,68,68,0.3); }
        .mb-warn { background: rgba(245,158,11,0.15); color: #f59e0b; border: 1px solid rgba(245,158,11,0.3); }

        .row-count { font-size: 0.82em; color: #94a3b8; }
        .arrow { color: #64748b; transition: transform 0.2s; }
        .arrow.open { transform: rotate(180deg); }

        .month-bar {
            padding: 8px 15px; font-size: 0.82em;
            display: flex; align-items: center; gap: 8px;
        }
        .month-bar.match { background: rgba(16,185,129,0.08); color: #6ee7b7; border-top: 1px solid rgba(16,185,129,0.15); }
        .month-bar.mismatch { background: rgba(239,68,68,0.08); color: #fca5a5; border-top: 1px solid rgba(239,68,68,0.15); }
        .month-bar.warn { background: rgba(245,158,11,0.08); color: #fde68a; border-top: 1px solid rgba(245,158,11,0.15); }
        .month-bar .detected-list { font-size: 0.9em; color: #94a3b8; margin-left: auto; }

        .parse-bar {
            padding: 6px 15px; font-size: 0.8em; color: #a78bfa;
            background: rgba(167,139,250,0.06);
            border-top: 1px solid rgba(167,139,250,0.12);
        }

        .preview-area {
            display: none; border-top: 1px solid #334155;
            max-height: 500px; overflow: auto; background: #0c1222;
        }
        .data-table {
            width: 100%; border-collapse: collapse; font-size: 0.8em;
            font-family: "Consolas", "Courier New", monospace;
        }
        .data-table th {
            background: #1a3a5c; color: #93c5fd; padding: 6px 10px;
            text-align: left; font-weight: 600; position: sticky; top: 0;
            z-index: 1; border-bottom: 2px solid #2563eb; white-space: nowrap;
        }
        .data-table td {
            padding: 5px 10px; border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #cbd5e1; white-space: nowrap;
        }
        .data-table tr:hover td { background: rgba(59, 130, 246, 0.08); }
        .data-table .row-num {
            color: #475569; text-align: right; padding-right: 12px;
            font-size: 0.85em; user-select: none;
            border-right: 1px solid #334155; background: #111827;
        }
        .data-table .cell-empty { color: #334155; font-style: italic; }

        .text-preview {
            padding: 15px; font-family: monospace; font-size: 0.8em;
            color: #10b981; white-space: pre-wrap; line-height: 1.6;
        }
        .text-preview .line-num {
            display: inline-block; width: 40px; color: #475569;
            text-align: right; margin-right: 12px; user-select: none;
        }
        .error-preview { padding: 15px; color: #ef4444; font-size: 0.85em; }
    </style>
</head>
<body>
<div class="container">
    <h1>📂 Tenant Report Reader</h1>
    <p class="subtitle">Upload tenant reports — the engine extracts sales data into a unified master report.</p>

    <div class="config-card">
        <div class="config-title">📅 Report Month</div>
        <div class="config-row">
            <div class="config-group">
                <label>Month</label>
                <select id="monthSel">
                    <option value="1">January</option><option value="2">February</option>
                    <option value="3">March</option><option value="4">April</option>
                    <option value="5">May</option><option value="6">June</option>
                    <option value="7">July</option><option value="8">August</option>
                    <option value="9">September</option><option value="10">October</option>
                    <option value="11">November</option><option value="12">December</option>
                </select>
            </div>
            <div class="config-group">
                <label>Year</label>
                <input id="yearIn" type="number" value="2026" min="2000" max="2100" style="width:90px"/>
            </div>
            <div class="month-preview">Target: <b id="targetLabel">January 2026</b></div>
        </div>
    </div>

    <div class="drop-zone" id="dropZone">
        <div class="icon">📥</div>
        <p><b>Drop files or folder here</b></p>
        <p style="font-size: 0.85em;">Or use the buttons below</p>
    </div>

    <div class="browse-row">
        <button class="browse-btn btn-files" onclick="document.getElementById('fileInput').click()">📄 Browse Files</button>
        <button class="browse-btn btn-folder" onclick="document.getElementById('folderInput').click()">📁 Browse Folder</button>
        <button class="browse-btn btn-clear" onclick="clearAll()">↺ Clear All</button>
    </div>

    <input type="file" id="fileInput" class="hidden-input" accept=".xlsx,.xls,.pdf" multiple />
    <input type="file" id="folderInput" class="hidden-input" webkitdirectory />

    <div id="loader" class="loader"><span class="spinner"></span> Reading, validating &amp; parsing files...</div>

    <div class="summary" id="summary">
        <div class="summary-grid">
            <div class="summary-card"><div class="num c-blue" id="sTotal">0</div><div class="label">Total Files</div></div>
            <div class="summary-card"><div class="num c-green" id="sOk">0</div><div class="label">Read OK</div></div>
            <div class="summary-card"><div class="num c-red" id="sFail">0</div><div class="label">Read Failed</div></div>
            <div class="summary-card"><div class="num c-purple" id="sMatch">0</div><div class="label">Month Match ✅</div></div>
            <div class="summary-card"><div class="num c-yellow" id="sWrong">0</div><div class="label">Wrong Month ⚠️</div></div>
        </div>
    </div>

    <!-- MASTER REPORT -->
    <div class="master" id="masterSection">
        <div class="config-card">
            <div class="config-title">📊 Master Report — Unified Tenant Sales</div>
            <div class="master-warn" id="masterWarn"></div>
            <div class="master-table-wrap">
                <table class="master-table" id="masterTable"></table>
            </div>
        </div>
    </div>

    <div class="file-list" id="fileList"></div>
</div>

<script>
var dropZone    = document.getElementById("dropZone");
var fileList    = document.getElementById("fileList");
var loader      = document.getElementById("loader");
var summaryEl   = document.getElementById("summary");
var fileInput   = document.getElementById("fileInput");
var folderInput = document.getElementById("folderInput");

function updateLabel() {
    var months = ["","January","February","March","April","May","June",
                  "July","August","September","October","November","December"];
    var m = parseInt(document.getElementById("monthSel").value);
    var y = document.getElementById("yearIn").value;
    document.getElementById("targetLabel").textContent = months[m] + " " + y;
}
document.getElementById("monthSel").addEventListener("change", updateLabel);
document.getElementById("yearIn").addEventListener("input", updateLabel);
var now = new Date();
document.getElementById("monthSel").value = now.getMonth() + 1;
document.getElementById("yearIn").value = now.getFullYear();
updateLabel();

dropZone.addEventListener("dragover", function(e) { e.preventDefault(); dropZone.classList.add("hover"); });
dropZone.addEventListener("dragleave", function() { dropZone.classList.remove("hover"); });
dropZone.addEventListener("drop", async function(e) {
    e.preventDefault(); dropZone.classList.remove("hover");
    var items = e.dataTransfer.items;
    var formData = new FormData();
    for (var i = 0; i < items.length; i++) {
        var entry = items[i].webkitGetAsEntry();
        if (entry) await walkTree(entry, formData, "");
    }
    sendFiles(formData);
});

fileInput.addEventListener("change", function() {
    var files = fileInput.files;
    if (!files.length) return;
    var formData = new FormData();
    for (var i = 0; i < files.length; i++) {
        var name = files[i].name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf"))
            formData.append("files", files[i], files[i].name);
    }
    fileInput.value = "";
    sendFiles(formData);
});

folderInput.addEventListener("change", function() {
    var files = folderInput.files;
    if (!files.length) return;
    var formData = new FormData();
    for (var i = 0; i < files.length; i++) {
        var name = files[i].name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf"))
            formData.append("files", files[i], files[i].webkitRelativePath || files[i].name);
    }
    folderInput.value = "";
    sendFiles(formData);
});

function walkTree(item, formData, path) {
    return new Promise(function(resolve) {
        if (item.isFile) {
            item.file(function(file) {
                var name = file.name.toLowerCase();
                if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf"))
                    formData.append("files", file, path + file.name);
                resolve();
            });
        } else if (item.isDirectory) {
            var reader = item.createReader();
            readAllEntries(reader, function(entries) {
                var promises = [];
                for (var i = 0; i < entries.length; i++)
                    promises.push(walkTree(entries[i], formData, path + item.name + "/"));
                Promise.all(promises).then(resolve);
            });
        } else { resolve(); }
    });
}

function readAllEntries(reader, callback) {
    var all = [];
    function batch() {
        reader.readEntries(function(entries) {
            if (entries.length === 0) callback(all);
            else { all = all.concat(Array.from(entries)); batch(); }
        });
    }
    batch();
}

async function sendFiles(formData) {
    fileList.innerHTML = "";
    summaryEl.style.display = "none";
    document.getElementById("masterSection").style.display = "none";
    loader.style.display = "block";

    formData.append("month", document.getElementById("monthSel").value);
    formData.append("year", document.getElementById("yearIn").value);

    try {
        var resp = await fetch("/upload", { method: "POST", body: formData });
        var data = await resp.json();
        loader.style.display = "none";
        renderAll(data);
    } catch (err) {
        loader.style.display = "none";
        alert("Error: " + err.message);
    }
}

function clearAll() {
    fileList.innerHTML = "";
    summaryEl.style.display = "none";
    document.getElementById("masterSection").style.display = "none";
    loader.style.display = "none";
}

function targetKey() {
    var m = parseInt(document.getElementById("monthSel").value);
    var y = parseInt(document.getElementById("yearIn").value);
    return y + "-" + (m < 10 ? "0" + m : "" + m);
}

function renderAll(data) {
    if (data.error) { alert(data.error); return; }
    var results = data.results;
    var total = results.length, ok = 0, fail = 0, matched = 0, wrong = 0;

    for (var i = 0; i < results.length; i++) {
        var r = results[i];
        if (r.success) ok++; else fail++;
        if (r.month_check) {
            if (r.month_check.status === "mismatch") wrong++;
            else if (r.month_check.status === "ok") matched++;
        }
    }

    document.getElementById("sTotal").textContent = total;
    document.getElementById("sOk").textContent = ok;
    document.getElementById("sFail").textContent = fail;
    document.getElementById("sMatch").textContent = matched;
    document.getElementById("sWrong").textContent = wrong;
    summaryEl.style.display = "block";

    results.sort(function(a, b) {
        var order = {"mismatch": 0, "warning": 1, "ok_multi": 2, "ok": 3};
        var sa = a.month_check ? (order[a.month_check.status] || 3) : 3;
        var sb = b.month_check ? (order[b.month_check.status] || 3) : 3;
        return sa - sb;
    });

    buildMaster(results);

    for (var i = 0; i < results.length; i++) renderCard(results[i], i);
}

// ─── MASTER REPORT ─────────────────────────────────────────

function buildMaster(results) {
    var tenantMap = {};
    var allMonths = {};
    var unparsed = [];

    results.forEach(function(res) {
        if (res.parsed && res.parsed.success) {
            var ts = res.parsed.tenants;
            Object.keys(ts).forEach(function(t) {
                if (!tenantMap[t]) tenantMap[t] = { monthly: {}, files: [], dailyCount: 0 };
                if (tenantMap[t].files.indexOf(res.filename) === -1)
                    tenantMap[t].files.push(res.filename);
                var m = ts[t].monthly || {};
                Object.keys(m).forEach(function(k) {
                    if (!(k in tenantMap[t].monthly)) tenantMap[t].monthly[k] = m[k];
                    allMonths[k] = true;
                });
                tenantMap[t].dailyCount += (ts[t].daily || []).length;
            });
        } else {
            var msg = res.filename;
            if (res.parsed && res.parsed.message) msg += " — " + res.parsed.message;
            unparsed.push(msg);
        }
    });

    var tenantNames = Object.keys(tenantMap);
    if (!tenantNames.length) {
        document.getElementById("masterSection").style.display = "none";
        return;
    }

    var months = Object.keys(allMonths).sort();
    var tKey = targetKey();

    // Sort tenants by target-month value descending
    tenantNames.sort(function(a, b) {
        return (tenantMap[b].monthly[tKey] || 0) - (tenantMap[a].monthly[tKey] || 0);
    });

    var html = "<thead><tr><th>Tenant</th>";
    months.forEach(function(mk) {
        var parts = mk.split("-");
        var label = monthShort(parseInt(parts[1])) + "-" + parts[0].slice(2);
        html += '<th class="' + (mk === tKey ? "t-head" : "") + '">' + label + '</th>';
    });
    html += "<th>Total</th></tr></thead><tbody>";

    var colTotals = {};
    var grand = 0;

    tenantNames.forEach(function(t) {
        var tm = tenantMap[t];
        var rowTotal = 0;
        html += "<tr><td><b>" + esc(t) + "</b><br><span class='src'>" +
                esc(tm.files.join(", ")) +
                (tm.dailyCount ? " · " + tm.dailyCount + " daily rows" : "") +
                "</span></td>";
        months.forEach(function(mk) {
            var v = tm.monthly[mk];
            if (v !== undefined) {
                rowTotal += v;
                colTotals[mk] = (colTotals[mk] || 0) + v;
                html += '<td class="' + (mk === tKey ? "t-col" : "") + '">' + fmtNum(v) + '</td>';
            } else {
                html += '<td class="no-data ' + (mk === tKey ? "t-col" : "") + '">—</td>';
            }
        });
        grand += rowTotal;
        html += "<td><b>" + fmtNum(rowTotal) + "</b></td></tr>";
    });

    html += '<tr class="total-row"><td>🏆 GRAND TOTAL</td>';
    months.forEach(function(mk) {
        html += '<td class="' + (mk === tKey ? "t-col" : "") + '">' + fmtNum(colTotals[mk] || 0) + '</td>';
    });
    html += "<td>" + fmtNum(grand) + "</td></tr>";
    html += "</tbody>";

    document.getElementById("masterTable").innerHTML = html;

    var warnEl = document.getElementById("masterWarn");
    if (unparsed.length) {
        warnEl.style.display = "block";
        warnEl.innerHTML = "⚠️ <b>" + unparsed.length + " file(s) could not be parsed:</b><br>• " +
            unparsed.map(esc).join("<br>• ");
    } else {
        warnEl.style.display = "none";
    }

    document.getElementById("masterSection").style.display = "block";
}

function monthShort(m) {
    return ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m];
}

// ─── FILE CARDS ────────────────────────────────────────────

function renderCard(res, idx) {
    var card = document.createElement("div");
    card.className = "file-card";

    if (res.month_check) {
        if (res.month_check.status === "mismatch") card.classList.add("mismatch");
        else if (res.month_check.status === "ok") card.classList.add("matched");
        else card.classList.add("warning");
    }

    var ext = res.filename.split(".").pop().toLowerCase();
    var extClass = ext === "pdf" ? "b-pdf" : ext === "xls" ? "b-xls" : "b-xlsx";
    var statusClass = res.success ? "b-ok" : "b-fail";
    var statusText  = res.success ? "OK" : "FAIL";
    var rowText = (res.total_rows || 0).toLocaleString() + " rows";
    var pvId = "pv_" + idx, arId = "ar_" + idx;

    var monthBadge = "";
    if (res.month_check) {
        var mc = res.month_check;
        var mbClass = mc.status === "ok" ? "mb-match" : mc.status === "mismatch" ? "mb-mismatch" : "mb-warn";
        var shortMsg = mc.message.length > 50 ? mc.message.substring(0, 47) + "..." : mc.message;
        monthBadge = '<span class="month-badge ' + mbClass + '">' + mc.icon + " " + esc(shortMsg) + '</span>';
    }

    var parseBadge = "";
    if (res.parsed) {
        if (res.parsed.success)
            parseBadge = '<span class="badge b-parse">🧩 ' + esc(res.parsed.format) + '</span>';
        else
            parseBadge = '<span class="badge b-fail">🧩 unparsed</span>';
    }

    var header = document.createElement("div");
    header.className = "file-header";
    header.setAttribute("onclick", "toggle('" + pvId + "','" + arId + "')");
    header.innerHTML =
        '<div class="file-info">' +
            '<span class="badge ' + extClass + '">' + ext.toUpperCase() + '</span>' +
            '<b>' + esc(res.filename) + '</b>' +
            '<span class="badge ' + statusClass + '">' + statusText + '</span>' +
            monthBadge + parseBadge +
        '</div>' +
        '<div style="display:flex;align-items:center;gap:10px">' +
            '<span class="row-count">' + rowText + '</span>' +
            '<span class="arrow" id="' + arId + '">▼</span>' +
        '</div>';
    card.appendChild(header);

    if (res.month_check) {
        var mc = res.month_check;
        var barClass = mc.match ? "match" : mc.status === "mismatch" ? "mismatch" : "warn";
        var detStr = mc.detected.length > 0 ? "Detected: " + mc.detected.join(", ") : "No months detected";
        var bar = document.createElement("div");
        bar.className = "month-bar " + barClass;
        bar.innerHTML = '<span>' + mc.icon + ' ' + esc(mc.message) + '</span>' +
            '<span class="detected-list">' + esc(detStr) + '</span>';
        card.appendChild(bar);
    }

    if (res.parsed && res.parsed.message) {
        var pb = document.createElement("div");
        pb.className = "parse-bar";
        pb.textContent = "🧩 " + res.parsed.message;
        card.appendChild(pb);
    }

    var preview = document.createElement("div");
    preview.className = "preview-area";
    preview.id = pvId;

    if (!res.success) {
        preview.innerHTML = '<div class="error-preview">' + esc(res.error_message || "Unknown error") + '</div>';
    } else if (res.data_type === "table") {
        preview.innerHTML = buildTable(res.data_rows, res.data_cols);
    } else if (res.data_type === "text") {
        preview.innerHTML = buildText(res.data_lines);
    }

    card.appendChild(preview);
    fileList.appendChild(card);
}

function buildTable(rows, numCols) {
    var html = '<table class="data-table"><thead><tr><th class="row-num">#</th>';
    for (var c = 0; c < numCols; c++) {
        var letter;
        if (c < 26) letter = String.fromCharCode(65 + c);
        else letter = String.fromCharCode(64 + Math.floor(c/26)) + String.fromCharCode(65 + (c % 26));
        html += '<th>' + letter + '</th>';
    }
    html += '</tr></thead><tbody>';
    for (var r = 0; r < rows.length; r++) {
        html += '<tr><td class="row-num">' + (r+1) + '</td>';
        for (var c = 0; c < numCols; c++) {
            var val = c < rows[r].length ? rows[r][c] : "";
            if (val === "" || val === "nan" || val === "None")
                html += '<td class="cell-empty">&mdash;</td>';
            else
                html += '<td>' + esc(val) + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
}

function buildText(lines) {
    var html = '<div class="text-preview">';
    for (var i = 0; i < lines.length; i++)
        html += '<span class="line-num">' + (i+1) + '</span>' + esc(lines[i]) + '\n';
    html += '</div>';
    return html;
}

function toggle(pvId, arId) {
    var el = document.getElementById(pvId);
    var ar = document.getElementById(arId);
    if (el.style.display === "block") { el.style.display = "none"; ar.classList.remove("open"); }
    else { el.style.display = "block"; ar.classList.add("open"); }
}

function esc(text) {
    var div = document.createElement("div");
    div.textContent = String(text);
    return div.innerHTML;
}

function fmtNum(v) { return Number(v).toLocaleString("id-ID"); }
</script>
</body>
</html>
"""


# ══════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template_string(HTML_UI)


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"error": "No files found"}), 400

    files = request.files.getlist("files")

    try:
        target_month = int(request.form.get("month", 1))
        target_year  = int(request.form.get("year", datetime.now().year))
    except ValueError:
        target_month = 1
        target_year  = datetime.now().year

    results = []

    for f in files:
        sid = str(uuid.uuid4())[:8]
        safe = sid + "_" + secure_filename(f.filename)
        filepath = UPLOAD_FOLDER / safe
        f.save(filepath)

        ext = filepath.suffix.lower()
        entry = {
            "filename": f.filename,
            "success": False,
            "total_rows": 0,
            "data_type": "error",
            "error_message": "",
            "month_check": None,
            "parsed": None,
        }

        if ext in [".xlsx", ".xls"]:
            data = read_excel(filepath)
        elif ext == ".pdf":
            data = read_pdf(filepath)
        else:
            data = {"type": "error", "message": "Unsupported type: " + ext}

        if data["type"] == "error":
            entry["error_message"] = data.get("message", "Unknown error")
        else:
            entry["success"] = True

            if data["type"] == "table":
                entry["data_type"]  = "table"
                entry["data_rows"]  = data["rows"]
                entry["data_cols"]  = data["cols"]
                entry["total_rows"] = len(data["rows"])
                detected = detect_months_in_excel(data["rows"])
            else:
                entry["data_type"]  = "text"
                entry["data_lines"] = data["lines"]
                entry["total_rows"] = len(data["lines"])
                detected = detect_months_in_lines(data["lines"])

            # Month detection also uses raw lines if available
            if data.get("lines"):
                detected.update(detect_months_in_lines(data["lines"]))
            if data.get("raw_lines"):
                detected.update(detect_months_in_lines(data["raw_lines"]))
            detected.update(detect_months_in_text(f.filename))
            entry["month_check"] = validate_month(detected, target_year, target_month)

            # ─── PHASE 2: Parse the report structure ────────────
            entry["parsed"] = parse_report(data, ext, f.filename)

        results.append(entry)

        if filepath.exists():
            os.remove(filepath)

    return jsonify({"results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Starting on port", port)
    app.run(host="0.0.0.0", port=port, debug=False)
