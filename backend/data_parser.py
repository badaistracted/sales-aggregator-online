# phase1_core_engine.py
"""
Phase 1: Core Engine & Data Parsing
- SMART SCANNER: finds headers and data anywhere in the file
- Handles merged cells, logo rows, title rows, empty rows
- Fuzzy column name matching (Date, Sales in any language/format)
- Works with .xlsx and .xls
- Indonesian holiday classification
"""

import re
import calendar
import warnings
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import holidays

# Suppress openpyxl warnings about styles in older files
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ══════════════════════════════════════════════════════════════════
#  SECTION 1 — Indonesian Holiday & Day Classification
# ══════════════════════════════════════════════════════════════════

def get_indonesian_holidays(year: int) -> set:
    """Return a set of date objects that are Indonesian national holidays."""
    return set(holidays.Indonesia(years=year).keys())


def classify_day(d: date, holiday_set: set) -> str:
    """
    Weekend  = Saturday, Sunday, OR a weekday that falls on a public holiday.
    Weekday  = everything else.
    """
    if d.weekday() >= 5:
        return "Weekend"
    if d in holiday_set:
        return "Weekend"
    return "Weekday"


def build_continuous_timeline(year: int, month: int) -> pd.DataFrame:
    """Every calendar day in the month with DayName and DayType."""
    holiday_set = get_indonesian_holidays(year)
    first_day   = date(year, month, 1)
    last_day    = date(year, month, calendar.monthrange(year, month)[1])
    days        = (last_day - first_day).days + 1

    records = []
    for i in range(days):
        d = first_day + timedelta(days=i)
        records.append({
            "Date"    : d,
            "DayName" : d.strftime("%A"),
            "DayType" : classify_day(d, holiday_set),
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
#  SECTION 2 — Fuzzy Column Name Matcher
# ══════════════════════════════════════════════════════════════════

# All reasonable aliases for "Date" and "Sales" columns
DATE_ALIASES = [
    "date", "tanggal", "tgl", "hari", "periode", "period",
    "transaction date", "trans date", "trx date", "trx_date",
    "trans_date", "transaction_date", "waktu", "datetime",
    "tgl transaksi", "tanggal transaksi",
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
    """Lowercase, strip whitespace, collapse spaces."""
    return re.sub(r"\s+", " ", str(text).lower().strip())


def _find_column(candidates: list[str], aliases: list[str]) -> str | None:
    """
    Given a list of column names from the DataFrame,
    return the first one that matches any alias.
    Tries exact match first, then 'starts with', then 'contains'.
    """
    norm_candidates = {_normalise(c): c for c in candidates}

    # Pass 1: exact match
    for alias in aliases:
        if alias in norm_candidates:
            return norm_candidates[alias]

    # Pass 2: candidate starts with alias
    for alias in aliases:
        for norm, original in norm_candidates.items():
            if norm.startswith(alias):
                return original

    # Pass 3: alias appears anywhere inside candidate
    for alias in aliases:
        for norm, original in norm_candidates.items():
            if alias in norm:
                return original

    # Pass 4: candidate appears inside any alias
    for alias in aliases:
        for norm, original in norm_candidates.items():
            if norm and norm in alias:
                return original

    return None


# ══════════════════════════════════════════════════════════════════
#  SECTION 3 — Smart File Scanner
# ══════════════════════════════════════════════════════════════════

class ScanResult:
    """Holds everything the scanner discovered about a file."""
    def __init__(self):
        self.success      : bool        = False
        self.header_row   : int         = -1   # 0-indexed row in the raw sheet
        self.date_col     : str         = ""
        self.sales_col    : str         = ""
        self.df           : pd.DataFrame | None = None
        self.raw_preview  : list[list]  = []   # first 15 rows for debugging
        self.error        : str         = ""
        self.warnings     : list[str]   = []


def _looks_like_date(value) -> bool:
    """Return True if the value can be interpreted as a date."""
    if pd.isna(value):
        return False
    if isinstance(value, (date,)):
        return True
    if hasattr(value, "date"):          # datetime
        return True
    if isinstance(value, (int, float)):
        # Excel serial dates are typically 40000-50000 range
        if 30000 < value < 60000:
            return True
    if isinstance(value, str):
        # Try common date patterns
        patterns = [
            r"\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}",   # 01/01/2025
            r"\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}",       # 2025-01-01
            r"\d{1,2}\s+\w+\s+\d{4}",                      # 1 January 2025
            r"\w+\s+\d{1,2},?\s+\d{4}",                    # January 1, 2025
        ]
        for p in patterns:
            if re.search(p, str(value).strip()):
                return True
    return False


def _looks_like_number(value) -> bool:
    """Return True if value looks like a sales figure."""
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        # Remove currency symbols and separators, then check
        cleaned = re.sub(r"[Rp,\.\s$€£]", "", str(value)).strip()
        try:
            float(cleaned)
            return True
        except ValueError:
            return False
    return False


def _score_row_as_header(row_values: list) -> float:
    """
    Score a row (0.0 to 1.0) on how likely it is to be a header row.
    Higher = more likely to be a header.
    """
    if not row_values:
        return 0.0

    non_null = [v for v in row_values if not pd.isna(v) and str(v).strip()]
    if not non_null:
        return 0.0

    score = 0.0
    text_count = 0

    for v in non_null:
        s = str(v).strip()

        # Headers are usually short strings
        if isinstance(v, str) and len(s) < 60:
            text_count += 1

        # Bonus if it matches known aliases
        norm = _normalise(s)
        if norm in DATE_ALIASES or any(a in norm for a in DATE_ALIASES):
            score += 2.0
        if norm in SALES_ALIASES or any(a in norm for a in SALES_ALIASES):
            score += 2.0

        # Penalty if it looks like a date (headers shouldn't be dates)
        if _looks_like_date(v):
            score -= 1.5

        # Penalty if it looks like a pure number
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            score -= 0.5

    # Bonus if most cells are text
    if non_null:
        score += (text_count / len(non_null)) * 1.5

    return score


def _read_raw_sheet(filepath: Path) -> pd.DataFrame:
    """
    Read the entire sheet as raw strings with no header assumption.
    Returns a DataFrame where row 0 = first sheet row.
    """
    ext = filepath.suffix.lower()

    if ext == ".xlsx":
        raw = pd.read_excel(
            filepath,
            header=None,        # <-- KEY: don't assume any row is a header
            dtype=str,          # read everything as string first
            engine="openpyxl",
        )
    elif ext == ".xls":
        raw = pd.read_excel(
            filepath,
            header=None,
            dtype=str,
            engine="xlrd",
        )
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    return raw


def _read_raw_sheet_typed(filepath: Path, header_row: int) -> pd.DataFrame:
    """
    Re-read the file with proper typing, now that we know which row is the header.
    """
    ext = filepath.suffix.lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"

    return pd.read_excel(
        filepath,
        header=header_row,
        engine=engine,
    )


def smart_scan(filepath: str | Path, year: int, month: int) -> ScanResult:
    """
    THE MAIN SCANNER.

    Strategy:
    1. Read the entire file as raw strings (no header assumption)
    2. Score every row as a potential header row
    3. Pick the best candidate
    4. Re-read the file with that row as the header
    5. Use fuzzy matching to find Date and Sales columns
    6. Extract and validate the data
    7. Return a detailed ScanResult
    """
    filepath  = Path(filepath)
    result    = ScanResult()

    # ── Step 1: Raw read ──────────────────────────────────────────────
    try:
        raw = _read_raw_sheet(filepath)
    except Exception as exc:
        result.error = f"Cannot open file: {exc}"
        return result

    if raw.empty:
        result.error = "File appears to be empty."
        return result

    # Save a preview for debugging (first 15 rows)
    result.raw_preview = raw.head(15).values.tolist()

    # ── Step 2: Score every row as a potential header ─────────────────
    max_scan_rows = min(50, len(raw))   # Only scan first 50 rows
    scores        = []

    for row_idx in range(max_scan_rows):
        row_values = raw.iloc[row_idx].tolist()
        score      = _score_row_as_header(row_values)
        scores.append((score, row_idx, row_values))

    # Sort by score descending
    scores.sort(key=lambda x: x[0], reverse=True)

    best_score, best_row_idx, best_row_values = scores[0]

    # ── Step 3: Sanity check — need a minimum score ───────────────────
    if best_score < 0.5:
        # No row looked like a header — try a fallback:
        # look for the first row that has date-like and number-like values
        result.warnings.append(
            f"No clear header row found (best score: {best_score:.2f}). "
            "Attempting data-row detection fallback."
        )
        header_row_idx = _fallback_find_data_start(raw, result)
        if header_row_idx == -1:
            result.error = (
                "Could not find headers or data in this file. "
                "Please ensure the file has a 'Date' column and a 'Sales' column."
            )
            return result
    else:
        header_row_idx = best_row_idx

    result.header_row = header_row_idx

    # ── Step 4: Re-read with proper header ───────────────────────────
    try:
        df = _read_raw_sheet_typed(filepath, header_row_idx)
    except Exception as exc:
        result.error = f"Failed to re-read file with header at row {header_row_idx}: {exc}"
        return result

    if df.empty:
        result.error = "No data found after the header row."
        return result

    # Drop columns that are entirely empty
    df.dropna(axis=1, how="all", inplace=True)
    # Drop rows that are entirely empty
    df.dropna(axis=0, how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # ── Step 5: Fuzzy column matching ────────────────────────────────
    col_names = [str(c) for c in df.columns]

    date_col  = _find_column(col_names, DATE_ALIASES)
    sales_col = _find_column(col_names, SALES_ALIASES)

    # If fuzzy match failed, try sniffing by content
    if not date_col:
        date_col = _sniff_date_column(df)
        if date_col:
            result.warnings.append(
                f"No date column name recognised. "
                f"Detected date column by content: '{date_col}'"
            )

    if not sales_col:
        sales_col = _sniff_numeric_column(df, exclude=date_col)
        if sales_col:
            result.warnings.append(
                f"No sales column name recognised. "
                f"Detected sales column by content: '{sales_col}'"
            )

    if not date_col:
        result.error = (
            f"Cannot find a Date column. "
            f"Columns found: {col_names}. "
            f"Expected a column named like: date, tanggal, tgl, transaction date, etc."
        )
        return result

    if not sales_col:
        result.error = (
            f"Cannot find a Sales column. "
            f"Columns found: {col_names}. "
            f"Expected a column named like: sales, penjualan, total, amount, revenue, etc."
        )
        return result

    result.date_col  = date_col
    result.sales_col = sales_col

    # ── Step 6: Parse and clean the data ─────────────────────────────
    df, parse_warnings = _parse_and_clean(df, date_col, sales_col, year, month)
    result.warnings.extend(parse_warnings)

    if df.empty:
        result.error = (
            f"No valid rows found for {month}/{year} after parsing. "
            f"Check that the file contains data for the selected month and year."
        )
        return result

    result.df      = df
    result.success = True
    return result


def _fallback_find_data_start(raw: pd.DataFrame, result: ScanResult) -> int:
    """
    Fallback: find the first row where at least one cell looks like a date
    AND at least one cell looks like a number. That row is probably the first
    data row — use the row BEFORE it as header (or row 0 if it's the first row).
    """
    for i in range(min(50, len(raw))):
        row = raw.iloc[i].tolist()
        has_date   = any(_looks_like_date(v)   for v in row)
        has_number = any(_looks_like_number(v) for v in row)
        if has_date and has_number:
            # Use the previous row as header if possible
            return max(0, i - 1)
    return -1


def _sniff_date_column(df: pd.DataFrame) -> str | None:
    """
    Scan each column's values to find one that contains dates.
    Returns the column name, or None.
    """
    for col in df.columns:
        sample = df[col].dropna().head(10).tolist()
        date_hits = sum(1 for v in sample if _looks_like_date(v))
        if date_hits >= max(1, len(sample) * 0.5):
            return str(col)
    return None


def _sniff_numeric_column(df: pd.DataFrame, exclude: str | None = None) -> str | None:
    """
    Find the column with the largest numeric values
    (assumes sales figures are bigger than, say, row counts or IDs).
    """
    best_col   = None
    best_mean  = 0

    for col in df.columns:
        if col == exclude:
            continue
        try:
            cleaned = (
                df[col]
                .astype(str)
                .str.replace(r"[Rp,\.\s$€£]", "", regex=True)
                .str.strip()
            )
            numeric = pd.to_numeric(cleaned, errors="coerce").dropna()
            if len(numeric) == 0:
                continue
            mean_val = numeric.mean()
            if mean_val > best_mean:
                best_mean = mean_val
                best_col  = str(col)
        except Exception:
            continue

    return best_col


def _parse_and_clean(
    df: pd.DataFrame,
    date_col: str,
    sales_col: str,
    year: int,
    month: int,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Parse dates, clean sales figures, filter to the target month.
    Returns (cleaned_df, list_of_warnings).
    """
    warnings_out = []
    work = df[[date_col, sales_col]].copy()
    work.columns = ["Date", "Sales"]

    # ── Parse dates ───────────────────────────────────────────────────
    original_count = len(work)

    # Try pandas auto-detection first
    work["Date"] = pd.to_datetime(work["Date"], errors="coerce", dayfirst=True)

    failed_dates = work["Date"].isna().sum()
    if failed_dates > 0:
        warnings_out.append(
            f"{failed_dates} rows had unparseable dates and were dropped."
        )

    work.dropna(subset=["Date"], inplace=True)
    work["Date"] = work["Date"].dt.date

    # ── Filter to target month ────────────────────────────────────────
    work = work[
        (work["Date"].apply(lambda d: d.year)  == year) &
        (work["Date"].apply(lambda d: d.month) == month)
    ]

    if len(work) == 0:
        return pd.DataFrame(), warnings_out

    # ── Clean sales values ────────────────────────────────────────────
    work["Sales"] = (
        work["Sales"]
        .astype(str)
        .str.replace(r"[Rp,\.\s$€£]", "", regex=True)
        .str.strip()
    )
    work["Sales"] = pd.to_numeric(work["Sales"], errors="coerce").fillna(0)
    work["Sales"] = work["Sales"].clip(lower=0)   # no negative sales

    # Remove duplicate dates (keep highest sales value — probably the summary row)
    dupes = work.duplicated(subset=["Date"], keep=False).sum()
    if dupes > 0:
        warnings_out.append(
            f"{dupes} duplicate date entries found. Keeping the row with the highest sales."
        )
        work = work.sort_values("Sales", ascending=False).drop_duplicates(
            subset=["Date"], keep="first"
        )

    work.sort_values("Date", inplace=True)
    work.reset_index(drop=True, inplace=True)

    return work, warnings_out


# ══════════════════════════════════════════════════════════════════
#  SECTION 4 — Multi-Tenant Aggregator
# ══════════════════════════════════════════════════════════════════

def parse_all_tenants(
    file_paths: list[str | Path],
    year: int,
    month: int,
) -> dict[str, pd.DataFrame]:
    """
    Scan every uploaded file, merge with the continuous timeline,
    and return { tenant_name: DataFrame[Date, DayName, DayType, Sales] }.
    """
    timeline  = build_continuous_timeline(year, month)
    tenants   = {}

    for fp in file_paths:
        fp   = Path(fp)
        name = fp.stem

        print(f"\n  Scanning: {fp.name}")
        print(f"  {'─' * 45}")

        result = smart_scan(fp, year, month)

        if not result.success:
            print(f"  ✗  FAILED: {result.error}")
            continue

        # Report what was found
        print(f"  ✓  Header found at row : {result.header_row + 1} "
              f"(Excel row {result.header_row + 2})")
        print(f"  ✓  Date column         : '{result.date_col}'")
        print(f"  ✓  Sales column        : '{result.sales_col}'")

        if result.warnings:
            for w in result.warnings:
                print(f"  ⚠  {w}")

        # Merge with full timeline (fills missing dates with Sales = 0)
        merged = timeline.merge(result.df, on="Date", how="left")
        merged["Sales"] = merged["Sales"].fillna(0)

        found_rows = result.df["Sales"].gt(0).sum()
        filled     = merged["Sales"].eq(0).sum()
        print(f"  ✓  Rows with sales data: {found_rows}")
        print(f"  ✓  Dates auto-filled   : {filled} (set to 0)")

        tenants[name] = merged

    return tenants


# ══════════════════════════════════════════════════════════════════
#  SECTION 5 — Console Printer (for validation)
# ══════════════════════════════════════════════════════════════════

def print_tenant_data(tenant_name: str, df: pd.DataFrame) -> None:
    width = 72
    print("═" * width)
    print(f"  TENANT : {tenant_name}")
    print("═" * width)
    print(f"  {'Date':<14} {'Day':<12} {'Type':<10} {'Sales (IDR)':>14}")
    print("  " + "─" * (width - 2))
    for _, row in df.iterrows():
        marker = "🟢" if row["DayType"] == "Weekend" else "  "
        print(
            f"  {str(row['Date']):<14}"
            f" {row['DayName']:<12}"
            f" {row['DayType']:<10}"
            f" {row['Sales']:>14,.0f}  {marker}"
        )
    print("  " + "─" * (width - 2))
    total  = df["Sales"].sum()
    wd_avg = df.loc[df["DayType"] == "Weekday",  "Sales"].mean()
    we_avg = df.loc[df["DayType"] == "Weekend",  "Sales"].mean()
    print(f"  {'Total Sales':<35} {total:>14,.0f}")
    print(f"  {'Weekday Avg / Day':<35} {wd_avg:>14,.0f}")
    print(f"  {'Weekend / Holiday Avg / Day':<35} {we_avg:>14,.0f}")
    print("═" * width)
    print()


# ══════════════════════════════════════════════════════════════════
#  SECTION 6 — Sample File Generator
# ══════════════════════════════════════════════════════════════════

def generate_sample_files(
    output_dir: str | Path = "sample_data",
    year: int = 2025,
    month: int = 1,
    tenants: list[str] | None = None,
) -> list[Path]:
    """
    Generate realistic messy Excel files that simulate real-world formats:
    - File A: Clean format (header at row 1)
    - File B: Has logo/title rows before the header
    - File C: Indonesian column names, extra metadata rows
    """
    if tenants is None:
        tenants = ["Tenant_A", "Tenant_B", "Tenant_C"]

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    rng       = np.random.default_rng(seed=42)
    first_day = date(year, month, 1)
    n_days    = calendar.monthrange(year, month)[1]
    all_dates = [first_day + timedelta(days=i) for i in range(n_days)]

    # Randomly drop ~20% of dates
    present   = [d for d in all_dates if rng.random() > 0.20]
    sales_all = rng.integers(500_000, 15_000_000, size=(len(tenants), len(present)))

    created = []
    for i, tenant in enumerate(tenants):
        out_path = output_dir / f"{tenant}.xlsx"
        sales    = sales_all[i]

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            if i == 0:
                # ── Clean format ──────────────────────────────────────
                df = pd.DataFrame({"Date": present, "Sales": sales})
                df.to_excel(writer, index=False, sheet_name="Sheet1")
                print(f"  Generated (clean format)     : {out_path.name}")

            elif i == 1:
                # ── Messy format: title rows before header ────────────
                # Write manually to simulate real-world mess
                wb = writer.book
                ws = wb.create_sheet("Sheet1")
                ws.append(["LAPORAN PENJUALAN HARIAN"])           # row 1: title
                ws.append(["PT. Contoh Tenant Indonesia"])        # row 2: company
                ws.append([])                                      # row 3: blank
                ws.append([f"Periode: {month}/{year}"])           # row 4: period
                ws.append([])                                      # row 5: blank
                ws.append(["Tanggal", "Jumlah Transaksi",        # row 6: HEADER
                            "Total Penjualan", "Keterangan"])
                for d, s in zip(present, sales):                  # rows 7+: data
                    ws.append([d.strftime("%d/%m/%Y"), rng.integers(10,100), s, "-"])
                # Remove default sheet
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]
                print(f"  Generated (messy Indonesian) : {out_path.name}")

            elif i == 2:
                # ── Mixed format: extra columns, different names ───────
                wb = writer.book
                ws = wb.create_sheet("Data")
                ws.append(["Store Report", None, None, None, None])
                ws.append([f"Store: {tenant}", None, None, None, None])
                ws.append([None])
                ws.append(["No", "Transaction Date", "Items Sold",  # row 4: HEADER
                            "Sales Amount", "Status"])
                for j, (d, s) in enumerate(zip(present, sales), start=1):
                    ws.append([j, d, rng.integers(5, 50), s, "Completed"])
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]
                print(f"  Generated (extra columns)    : {out_path.name}")

        created.append(out_path)

    return created


# ══════════════════════════════════════════════════════════════════
#  SECTION 7 — Entry Point
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    TARGET_YEAR  = 2025
    TARGET_MONTH = 1

    print("\n── Generating MESSY sample files ───────────────────────────────")
    sample_files = generate_sample_files(
        year=TARGET_YEAR,
        month=TARGET_MONTH,
    )

    print("\n── Smart-scanning all files ────────────────────────────────────")
    tenants_data = parse_all_tenants(sample_files, TARGET_YEAR, TARGET_MONTH)

    print("\n── Validation Output ───────────────────────────────────────────\n")
    for name, df in tenants_data.items():
        print_tenant_data(name, df)
