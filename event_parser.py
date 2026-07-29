# event_parser.py
"""
Parser for the mall's internal event calendar Excel file.

Input format:
  - Row with "PERIODE:" that indicates the month
  - Header row: Date | Time | Main Atrium | Amphitheatre | Foodtainment | Other | Category | Status
  - Dates can span multiple rows (merged cells in Excel -> empty cells in pandas)
  - Events are filled under location columns

Output:
  {
    "success": True/False,
    "daily": [...],           # Solution 1 format — list of {date, events: [...]}
    "monthly": {...},         # Solution 1 format — grouped by month
    "events_flat": [...],     # Flat list for easy lookup
    "is_events": True,        # Flag so app.py knows this is events data
    "message": "..."
  }
"""
import re
import pandas as pd
from datetime import datetime, date
from pathlib import Path

# Known month names for parsing the PERIODE header
MONTH_MAP = {
    "jan": 1, "januari": 1, "january": 1,
    "feb": 2, "februari": 2, "february": 2,
    "mar": 3, "maret": 3, "march": 3,
    "apr": 4, "april": 4,
    "mei": 5, "may": 5,
    "jun": 6, "juni": 6, "june": 6,
    "jul": 7, "juli": 7, "july": 7,
    "agu": 8, "ags": 8, "agustus": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "okt": 10, "oktober": 10, "october": 10,
    "nov": 11, "nop": 11, "nopember": 11, "november": 11,
    "des": 12, "desember": 12, "december": 12,
}

# Location columns we expect to find — ordered by likelihood
LOCATION_KEYS = [
    "main atrium",
    "amphitheatre",
    "foodtainment",
    "other",
]


def _clean_event_text(text):
    """Strip whitespace, newlines, and quotes from event text."""
    if text is None:
        return None
    s = str(text).strip()
    # Remove Excel-style wrapping quotes
    s = s.strip('"').strip("'")
    # Collapse multiple spaces and newlines into single space
    s = re.sub(r"\s+", " ", s)
    # Skip rows that are empty or just noise
    if s.lower() in ("", "nan", "none", "by eo"):
        return None
    # Skip rows that are only "by EO" variants
    if re.fullmatch(r"by\s+eo\b", s, re.IGNORECASE):
        return None
    return s


def _normalize_header(val):
    """Lowercase, strip, collapse spaces."""
    return re.sub(r"\s+", " ", str(val).strip().lower())


def _parse_date(val, year_hint=2026):
    """Parse date from Excel cell. Can be datetime, string, or float."""
    if val is None:
        return None

    # Already a date/datetime object
    if isinstance(val, (date, datetime)):
        d = val if isinstance(val, date) else val.date()
        return d

    # pandas Timestamp
    if hasattr(val, "date") and callable(val.date):
        try:
            return val.date()
        except Exception:
            pass

    s = str(val).strip()

    if not s or s.lower() in ("nan", "none", ""):
        return None

    # Try "DD/MM/YYYY" or "DD-MM-YYYY" or "YYYY-MM-DD"
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d",
                 "%d/%m/%y", "%d-%m-%y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # Excel serial number (float like 46169.0)
    try:
        f = float(s)
        from datetime import timedelta
        base = date(1899, 12, 30)
        return base + timedelta(days=int(f))
    except (ValueError, OverflowError):
        pass

    # "1/5/2026" without leading zeros — manual parse
    m = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})", s)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        if 1 <= day <= 31 and 1 <= month <= 12 and 2020 <= year <= 2035:
            try:
                return date(year, month, day)
            except ValueError:
                pass

    return None


def _parse_periode(periode_text):
    """Extract (year, month) from 'MEI 2026'."""
    s = str(periode_text).strip().lower()
    m = re.match(r"([a-z]+)\s+(\d{4})", s)
    if m:
        month_name = m.group(1)
        year = int(m.group(2))
        month_num = MONTH_MAP.get(month_name)
        if month_num and 2020 <= year <= 2035:
            return year, month_num
    return None, None


def try_event_parser(rows, filename=""):
    """
    Attempt to parse rows as a mall event calendar.

    Called from app.py's parse_report dispatcher alongside other parsers.

    Args:
        rows: list of lists from Excel read
        filename: original filename (unused, but kept for interface consistency)

    Returns:
        dict with success, daily, monthly, events_flat, is_events flag
    """
    # ── Step 1: Find the PERIODE row and header row ─────────────
    year_hint = None
    header_row_idx = None

    for idx, row in enumerate(rows[:30]):
        row_text = " ".join(str(c) for c in row).lower()
        if "periode" in row_text:
            for cell in row:
                parsed = _parse_periode(cell)
                if parsed[0]:
                    year_hint = parsed[0]
                    break

        normed = [_normalize_header(c) for c in row]
        # Header contains "date" AND at least one location keyword
        if any("date" in n or "tanggal" in n for n in normed):
            if any(loc in n for loc in LOCATION_KEYS for n in normed):
                header_row_idx = idx
                break

    if header_row_idx is None:
        return None

    if year_hint is None:
        year_hint = 2026

    header_normed = [_normalize_header(c) for c in rows[header_row_idx]]

    # ── Step 2: Map columns ──────────────────────────────────────
    col_map = {}  # location_name -> column index
    for name in LOCATION_KEYS:
        for c_idx, h in enumerate(header_normed):
            if name in h:
                col_map[name] = c_idx
                break

    date_col = None
    category_col = None
    status_col = None

    for c_idx, h in enumerate(header_normed):
        if "date" in h or "tanggal" in h:
            date_col = c_idx
        elif "categor" in h:
            category_col = c_idx
        elif "status" in h or "state" in h:
            status_col = c_idx

    if date_col is None:
        return None

    if not col_map:
        return None  # No location columns found — probably not an event file

    # ── Step 3: Extract events row by row ───────────────────────
    events_flat = []
    current_date = None

    for row in rows[header_row_idx + 1:]:
        if date_col >= len(row):
            continue

        parsed_date = _parse_date(row[date_col], year_hint)
        if parsed_date:
            current_date = parsed_date

        if current_date is None:
            continue

        date_str = str(current_date)

        for loc_name, loc_col in col_map.items():
            if loc_col >= len(row):
                continue
            event_text = _clean_event_text(row[loc_col])
            if not event_text:
                continue

            cat = ""
            sts = ""
            if category_col is not None and category_col < len(row):
                cat_raw = str(row[category_col]).strip()
                if cat_raw.lower() not in ("nan", "none", ""):
                    cat = cat_raw
            if status_col is not None and status_col < len(row):
                sts_raw = str(row[status_col]).strip()
                if sts_raw.lower() not in ("nan", "none", ""):
                    sts = sts_raw

            events_flat.append({
                "date": date_str,
                "event_name": event_text,
                "location": loc_name.title() if loc_name != "other" else "Other",
                "category": cat,
                "status": sts,
            })

    if len(events_flat) < 2:
        return None  # Too few events — probably not a real event file

    # ── Step 4: Deduplicate ──────────────────────────────────────
    seen = set()
    unique = []
    for e in events_flat:
        key = (e["date"], e["event_name"], e["location"])
        if key not in seen:
            seen.add(key)
            unique.append(e)

    unique.sort(key=lambda e: e["date"])

    # ── Step 5: Build Solution 1 format ──────────────────────────
    # daily: [{"date": "2026-05-01", "events": ["Event A", "Event B"]}, ...]
    from collections import defaultdict
    daily_map = defaultdict(list)
    for e in unique:
        daily_map[e["date"]].append(e["event_name"])

    daily = [
        {
            "date": d,
            "events": evts,
            "event_count": len(evts),
        }
        for d, evts in sorted(daily_map.items())
    ]

    # monthly: {"2026-05": [list of events for that month], ...}
    monthly_map = defaultdict(list)
    for e in unique:
        key = e["date"][:7]  # "2026-05"
        monthly_map[key].append(e)

    monthly = {
        k: {
            "events": v,
            "event_count": len(v),
            "event_days": len(set(e["date"] for e in v)),
        }
        for k, v in monthly_map.items()
    }

    # ── Step 6: Build message ─────────────────────────────────────
    event_dates = sorted(daily_map.keys())

    return {
        "success": True,
        "format": "events",
        "is_events": True,
        "daily": daily,
        "monthly": monthly,
        "events_flat": unique,
        "message": (
            f"Events: {len(unique)} event(s) across {len(event_dates)} day(s). "
            f"Range: {event_dates[0]} to {event_dates[-1]}."
            if event_dates else "No events found."
        ),
    }
