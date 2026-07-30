# event_parser.py
"""
Parser for the mall's internal event calendar Excel file.

Handles:
  - Multiple sheets (one per month)
  - Split two-row headers (Date on row A, locations on row B)
  - Merged date cells (dates spanning multiple rows)
  - Indonesian date formats (DD/MM/YYYY, day-first)
  - Pandas auto-converted datetime objects
"""
import re
import pandas as pd
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict

# Known month names
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

LOCATION_KEYS = [
    "main atrium",
    "amphitheatre",
    "foodtainment",
    "other",
]


def _clean_event_text(text):
    if text is None:
        return None
    # Don't process date/datetime objects as event text
    if isinstance(text, (date, datetime)):
        return None
    s = str(text).strip()
    s = s.strip('"').strip("'")
    s = re.sub(r"\s+", " ", s)
    if s.lower() in ("", "nan", "none", "by eo"):
        return None
    if re.fullmatch(r"by\s+eo\b", s, re.IGNORECASE):
        return None
    return s


def _normalize_header(val):
    if isinstance(val, (date, datetime)):
        return ""  # Don't treat dates as header text
    return re.sub(r"\s+", " ", str(val).strip().lower())


def _parse_date(val, year_hint=2026):
    if val is None:
        return None

    # Already a date/datetime object (pandas auto-converts Excel dates)
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val

    # pandas Timestamp
    if hasattr(val, "date") and callable(val.date):
        try:
            return val.date()
        except Exception:
            pass

    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None

    # "2026-05-01 00:00:00" or "2026-05-01" (from pandas str conversion)
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # DD/MM/YYYY (Indonesian convention, day first)
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # Excel serial number (float like 46169.0)
    try:
        f = float(s)
        if 40000 < f < 60000:
            base = date(1899, 12, 30)
            return base + timedelta(days=int(f))
    except (ValueError, OverflowError):
        pass

    # Manual regex for "1/5/2026" without leading zeros
    m = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})", s)
    if m:
        a = int(m.group(1))
        b = int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        if 2020 <= year <= 2035:
            # Day/month/year first (Indonesian)
            if 1 <= a <= 31 and 1 <= b <= 12:
                try:
                    return date(year, b, a)
                except ValueError:
                    pass
            # Fallback month/day/year
            if 1 <= a <= 12 and 1 <= b <= 31:
                try:
                    return date(year, a, b)
                except ValueError:
                    pass

    return None


def _parse_periode(text):
    if isinstance(text, (date, datetime)):
        return None, None
    s = str(text).strip().lower()
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
    Parse a SINGLE sheet's rows as event calendar.
    Handles split two-row headers.
    """
    # ── Step 1: Find PERIODE row and "Date" header row ──────────
    year_hint = None
    date_header_idx = None

    for idx, row in enumerate(rows[:30]):
        # Check for PERIODE
        row_text_parts = []
        for c in row:
            if not isinstance(c, (date, datetime)):
                row_text_parts.append(str(c))
        row_text = " ".join(row_text_parts).lower()

        if "periode" in row_text:
            for cell in row:
                parsed = _parse_periode(cell)
                if parsed[0]:
                    year_hint = parsed[0]
                    break

        # Skip rows that have actual datetime values (data rows)
        has_datetime = any(isinstance(c, (date, datetime)) for c in row)
        if has_datetime:
            continue

        normed = [_normalize_header(c) for c in row]
        if any("date" in n or "tanggal" in n for n in normed):
            if date_header_idx is None:
                date_header_idx = idx

    if date_header_idx is None:
        return None

    if year_hint is None:
        year_hint = 2026

    # ── Step 2: Map columns (check row A and row B) ─────────────
    row_a = [_normalize_header(c) for c in rows[date_header_idx]]
    row_b = (
        [_normalize_header(c) for c in rows[date_header_idx + 1]]
        if date_header_idx + 1 < len(rows)
        else []
    )

    col_map = {}
    for name in LOCATION_KEYS:
        found = False
        for c_idx, h in enumerate(row_a):
            if name in h:
                col_map[name] = c_idx
                found = True
                break
        if not found:
            for c_idx, h in enumerate(row_b):
                if name in h:
                    col_map[name] = c_idx
                    break

    # Date / category / status from row A
    date_col = None
    category_col = None
    status_col = None
    for c_idx, h in enumerate(row_a):
        if ("date" in h or "tanggal" in h) and date_col is None:
            date_col = c_idx
        elif "categor" in h:
            category_col = c_idx
        elif "status" in h or "state" in h:
            status_col = c_idx

    if date_col is None:
        return None
    if not col_map:
        return None

    # Figure out where data starts
    locations_on_row_b = any(
        name in row_b[col_map[name]]
        for name in col_map
        if col_map[name] < len(row_b)
    )
    header_row_idx = date_header_idx + 1 if locations_on_row_b else date_header_idx

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
                cat_raw = row[category_col]
                if not isinstance(cat_raw, (date, datetime)):
                    cat_raw = str(cat_raw).strip()
                    if cat_raw.lower() not in ("nan", "none", ""):
                        cat = cat_raw
            if status_col is not None and status_col < len(row):
                sts_raw = row[status_col]
                if not isinstance(sts_raw, (date, datetime)):
                    sts_raw = str(sts_raw).strip()
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
        return None

    # ── Step 4: Deduplicate ──────────────────────────────────────
    seen = set()
    unique = []
    for e in events_flat:
        key = (e["date"], e["event_name"], e["location"])
        if key not in seen:
            seen.add(key)
            unique.append(e)
    unique.sort(key=lambda e: e["date"])

    # ── Step 5: Build output format ──────────────────────────────
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

    monthly_map = defaultdict(list)
    for e in unique:
        key = e["date"][:7]
        monthly_map[key].append(e)

    monthly = {
        k: {
            "events": v,
            "event_count": len(v),
            "event_days": len(set(e["date"] for e in v)),
        }
        for k, v in monthly_map.items()
    }

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


# ── Multi-sheet reader ────────────────────────────────────────

def parse_event_file(filepath):
    """
    Read an event calendar Excel file with multiple tabs (one per month).
    Tries every sheet and merges all events together.
    """
    filepath = Path(filepath)

    try:
        engine = "xlrd" if filepath.suffix == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        sheet_names = xl.sheet_names
    except Exception:
        return None

    all_events_flat = []
    messages = []

    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None).fillna("")
            rows = []
            for row in df.values.tolist():
                converted = []
                for c in row:
                    # Keep date/datetime objects as-is for _parse_date
                    if isinstance(c, (date, datetime)):
                        converted.append(c)
                    elif hasattr(c, "date") and callable(c.date):
                        # pandas Timestamp — convert to python datetime
                        try:
                            converted.append(c.to_pydatetime())
                        except Exception:
                            converted.append(str(c) if str(c) != "" else "")
                    else:
                        converted.append(str(c) if str(c) != "" else "")
                rows.append(converted)
        except Exception:
            continue

        if not rows:
            continue

        result = try_event_parser(rows, sheet_name)
        if result and result.get("success"):
            all_events_flat.extend(result.get("events_flat", []))
            messages.append(f"Sheet '{sheet_name}': {result['message']}")

    if not all_events_flat:
        return None

    # Deduplicate
    seen = set()
    unique_flat = []
    for e in all_events_flat:
        key = (e["date"], e["event_name"], e["location"])
        if key not in seen:
            seen.add(key)
            unique_flat.append(e)
    unique_flat.sort(key=lambda e: e["date"])

    # Build daily
    daily_map = defaultdict(list)
    for e in unique_flat:
        daily_map[e["date"]].append(e["event_name"])

    daily = [
        {
            "date": d,
            "events": list(set(evts)),
            "event_count": len(set(evts)),
        }
        for d, evts in sorted(daily_map.items())
    ]

    # Build monthly
    monthly_map = defaultdict(list)
    for e in unique_flat:
        key = e["date"][:7]
        monthly_map[key].append(e)

    monthly = {
        k: {
            "events": v,
            "event_count": len(v),
            "event_days": len(set(e["date"] for e in v)),
        }
        for k, v in monthly_map.items()
    }

    months_found = sorted(monthly.keys())

    return {
        "success": True,
        "format": "events",
        "is_events": True,
        "daily": daily,
        "monthly": monthly,
        "events_flat": unique_flat,
        "message": (
            f"Event calendar: {len(unique_flat)} event(s) across "
            f"{len(months_found)} month(s) from {len(messages)} sheet(s). "
            f"Months: {', '.join(months_found)}"
        ),
    }
