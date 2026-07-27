"""
Phase 1: Core Engine & Data Parsing
------------------------------------
Reads tenant Excel files (each with a Date column and a Sales column),
builds a continuous daily timeline for the covered month (filling any
missing dates with 0 sales), and classifies every day as Weekday or
Weekend -- where Weekend also absorbs any Indonesian national holiday
that falls on a weekday.
"""
import calendar
import datetime as dt
from dataclasses import dataclass

import holidays
import pandas as pd

# Column names we accept, case/whitespace-insensitive, so tenants that
# label things slightly differently ("date", " Sales ", "SALES") still work.
DATE_ALIASES = {"date", "tanggal"}
SALES_ALIASES = {"sales", "sale", "penjualan", "total sales", "revenue"}


class TenantFileError(ValueError):
    """Raised when a tenant file is missing a required column or has no usable rows."""


@dataclass
class TenantData:
    tenant_name: str
    df: pd.DataFrame  # columns: Date, Day Name, Day Type, Sales
    year: int
    month: int


def _find_column(columns, aliases):
    for col in columns:
        if str(col).strip().lower() in aliases:
            return col
    return None


def _tenant_name_from_filename(filepath: str) -> str:
    import os

    base = os.path.basename(filepath)
    name = os.path.splitext(base)[0]
    return name.replace("_", " ").replace("-", " ").strip()


def read_tenant_file(filepath: str, tenant_name: str | None = None) -> pd.DataFrame:
    """Read a single tenant Excel file and return a raw (Date, Sales) frame.

    Raises TenantFileError if the required columns can't be found.
    """
    raw = pd.read_excel(filepath)

    date_col = _find_column(raw.columns, DATE_ALIASES)
    sales_col = _find_column(raw.columns, SALES_ALIASES)

    if date_col is None or sales_col is None:
        missing = []
        if date_col is None:
            missing.append("Date")
        if sales_col is None:
            missing.append("Sales")
        raise TenantFileError(
            f"Missing required column(s): {', '.join(missing)}. "
            f"Found columns: {list(raw.columns)}"
        )

    frame = raw[[date_col, sales_col]].copy()
    frame.columns = ["Date", "Sales"]
    frame["Date"] = pd.to_datetime(frame["Date"], errors="coerce")
    frame["Sales"] = pd.to_numeric(frame["Sales"], errors="coerce")

    frame = frame.dropna(subset=["Date"])
    if frame.empty:
        raise TenantFileError("No valid rows with a parseable Date were found.")

    # If a date appears more than once, sum same-day sales together.
    frame = frame.groupby("Date", as_index=False)["Sales"].sum()

    return frame


def build_continuous_month(frame: pd.DataFrame, id_holidays: "holidays.HolidayBase") -> pd.DataFrame:
    """Reindex a (Date, Sales) frame onto every day of the month it covers,
    filling missing dates with 0 sales, and classify each day.
    """
    min_date = frame["Date"].min()
    year, month = min_date.year, min_date.month
    days_in_month = calendar.monthrange(year, month)[1]
    full_range = pd.date_range(
        start=dt.date(year, month, 1), periods=days_in_month, freq="D"
    )

    frame = frame.set_index("Date").reindex(full_range).rename_axis("Date").reset_index()
    frame["Sales"] = frame["Sales"].fillna(0.0)

    frame["Day Name"] = frame["Date"].dt.day_name()

    def classify(d: pd.Timestamp) -> str:
        if d.weekday() >= 5:  # Saturday=5, Sunday=6
            return "Weekend"
        if d.date() in id_holidays:
            return "Weekend"  # holiday override
        return "Weekday"

    frame["Day Type"] = frame["Date"].apply(classify)

    return frame[["Date", "Day Name", "Day Type", "Sales"]]


def load_tenant(filepath: str, tenant_name: str | None = None) -> TenantData:
    tenant_name = tenant_name or _tenant_name_from_filename(filepath)
    raw = read_tenant_file(filepath, tenant_name)
    min_date = raw["Date"].min()
    id_holidays = holidays.country_holidays("ID", years=min_date.year)
    df = build_continuous_month(raw, id_holidays)
    return TenantData(tenant_name=tenant_name, df=df, year=min_date.year, month=min_date.month)


def load_tenants(filepaths: list[str]) -> tuple[list[TenantData], list[dict]]:
    """Load multiple tenant files. Returns (successes, errors) so the caller
    (web layer) can report per-file errors without aborting the whole batch.
    """
    successes: list[TenantData] = []
    errors: list[dict] = []
    for fp in filepaths:
        name = _tenant_name_from_filename(fp)
        try:
            successes.append(load_tenant(fp, tenant_name=name))
        except TenantFileError as exc:
            errors.append({"file": name, "message": str(exc)})
        except Exception as exc:  # pragma: no cover - defensive
            errors.append({"file": name, "message": f"Could not read file: {exc}"})
    return successes, errors


if __name__ == "__main__":
    # Console validation script (Phase 1 deliverable).
    import sys

    if len(sys.argv) < 2:
        print("Usage: python data_parser.py <tenant_file.xlsx> [more_files...]")
        sys.exit(1)

    for path in sys.argv[1:]:
        data = load_tenant(path)
        print(f"\n=== {data.tenant_name} ({calendar.month_name[data.month]} {data.year}) ===")
        print(data.df.to_string(index=False, formatters={"Date": lambda d: d.strftime("%Y-%m-%d")}))
