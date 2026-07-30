# event_parser.py
"""
Parser for the mall's internal event calendar Excel file.

Handles:
  - Multiple sheets (one per month)
  - Split two-row headers: Row A has Date/Time/Location/Category/Status
    Row B has sub-locations: Main Atrium / Amphitheatre / Foodtainment / Other
  - Merged date cells (dates spanning multiple rows)
  - Indonesian date formats
  - Pandas auto-converted datetime objects
"""
import re
import pandas as pd
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict

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

# All possible location sub-header names we might encounter
KNOWN_LOCATIONS = [
    "main atrium", "amphitheatre", "amphitheater",
    "foodtainment", "other",
]


def _is_date_obj(val):
    """Check if value is any kind of date/datetime object."""
    if isinstance(val, (date, datetime)):
        return True
    if hasattr(val, "date") and callable(val.date):
        return True
    return False


def _clean_event_text(text):
    if text is None:
        return None
    if _is_date_obj(text):
        return None
    s = str(text).strip()
    s = s.strip('"').strip("'")
    s = re.sub(r"\s+", " ", s)
    if s.lower() in ("", "nan", "none", "by eo", "true", "false",
                      "on progres", "on progress", "done", "cancel",
                      "cancelled", "pending"):
        return None
    if re.fullmatch(r"by\s+eo\b", s, re.IGNORECASE):
        return None
    # Skip if it's just a number
    try:
        float(s.replace(",", "").replace(".", ""))
        return None
    except ValueError:
        pass
    return s


def _norm(val):
    """Normalize cell for header matching. Returns '' for dates."""
    if _is_date_obj(val):
        return ""
    return re.sub(r"\s+", " ", str(val).strip().lower())


def _parse_date(val, year_hint=2026):
    if val is None:
        return None

    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val

    if hasattr(val, "date") and callable(val.date):
        try:
            return val.date()
        except Exception:
            pass

    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None

    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    try:
        f = float(s)
        if 40000 < f < 60000:
            base = date(1899, 12, 30)
            return base + timedelta(days=int(f))
    except (ValueError, OverflowError):
        pass

    m = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        if 2020 <= year <= 2035:
            if 1 <= a <= 31 and 1 <= b <= 12:
                try:
                    return date(year, b, a)
                except ValueError:
                    pass
            if 1 <= a <= 12 and 1 <= b <= 31:
                try:
                    return date(year, a, b)
                except ValueError:
                    pass

    return None


def _parse_periode(text):
    if _is_date_obj(text):
        return None, None
    s = str(text).strip().lower()
    # "MEI 2026" or "PERIODE: MEI 2026"
    m = re.search(r"([a-z]{3,})\s+(\d{4})", s)
    if m:
        month_name = m.group(1)
        year = int(m.group(2))
        month_num = MONTH_MAP.get(month_name)
        if month_num and 2020 <= year <= 2035:
            return year, month_num
    return None, None


def try_event_parser(rows, sheet_name=""):
    """
    Parse a SINGLE sheet's rows as event calendar.

    Strategy:
    1. Find the PERIODE row to know expected month/year
    2. Find the header row(s) — "Date" anchor + location sub-headers
    3. Extract events from each location column per day
    4. FILTER: only keep events matching the expected month
    """
    if not rows or len(rows) < 5:
        return None

    # ── Step 1: Find PERIODE and expected month ─────────────────
    expected_year = None
    expected_month = None

    for idx, row in enumerate(rows[:15]):
        for cell in row:
            if _is_date_obj(cell):
                continue
            yr, mn = _parse_periode(cell)
            if yr:
                expected_year = yr
                expected_month = mn
                break
        if expected_year:
            break

    # Also try sheet name: "MEI 2026", "JUNI 2026"
    if not expected_year and sheet_name:
        yr, mn = _parse_periode(sheet_name)
        if yr:
            expected_year = yr
            expected_month = mn

    # ── Step 2: Find the header structure ───────────────────────
    # Look for a row containing "Date" (or "Tanggal")
    date_header_idx = None

    for idx, row in enumerate(rows[:20]):
        # Skip rows with actual datetime values
        if any(_is_date_obj(c) for c in row):
            continue

        normed = [_norm(c) for c in row]

        has_date = any(
            n in ("date", "tanggal", "tgl", "hari/tanggal",
                  "hari/ tanggal", "hari / tanggal")
            or n.startswith("date") or n.startswith("tanggal")
            for n in normed
        )

        if has_date:
            date_header_idx = idx
            break

    if date_header_idx is None:
        return None

    # ── Step 3: Map columns from header rows ────────────────────
    # Row A = date_header_idx (has Date, Time, Location, Category, Status)
    # Row B = date_header_idx + 1 (has Main Atrium, Amphitheatre, etc.)
    row_a = rows[date_header_idx]
    row_a_norm = [_norm(c) for c in row_a]

    row_b = rows[date_header_idx + 1] if date_header_idx + 1 < len(rows) else []
    row_b_norm = [_norm(c) for c in row_b] if row_b else []

    # Find Date column in row A
    date_col = None
    for c_idx, n in enumerate(row_a_norm):
        if n in ("date", "tanggal", "tgl", "hari/tanggal",
                 "hari/ tanggal", "hari / tanggal") or \
           n.startswith("date") or n.startswith("tanggal"):
            date_col = c_idx
            break

    if date_col is None:
        return None

    # Find Category and Status columns in row A
    category_col = None
    status_col = None
    for c_idx, n in enumerate(row_a_norm):
        if "categor" in n or "kategori" in n:
            category_col = c_idx
        elif "status" in n:
            status_col = c_idx

    # Find "Location" in row A to determine the location column range
    location_start = None
    for c_idx, n in enumerate(row_a_norm):
        if "location" in n or "lokasi" in n or "tempat" in n:
            location_start = c_idx
            break

    # Map location sub-columns from row B
    # These are the columns between "Location" start and "Category" start
    location_end = category_col if category_col else len(row_b_norm)

    col_map = {}  # column_index -> location_name

    if row_b_norm:
        # Check row B for location sub-headers
        search_start = location_start if location_start is not None else 0
        for c_idx in range(search_start, min(location_end or len(row_b_norm), len(row_b_norm))):
            n = row_b_norm[c_idx]
            if not n:
                continue
            # Match against known location names
            matched = False
            for loc in KNOWN_LOCATIONS:
                if loc in n or n in loc:
                    col_map[c_idx] = loc.title()
                    matched = True
                    break
            if not matched and n not in ("nan", "none", ""):
                # Unknown location — still use it
                col_map[c_idx] = n.title()

    # If no locations found in row B, try row A itself
    if not col_map:
        for c_idx, n in enumerate(row_a_norm):
            for loc in KNOWN_LOCATIONS:
                if loc in n:
                    col_map[c_idx] = loc.title()
                    break

    if not col_map:
        return None

    # Determine where data rows start
    # If row B had location headers, data starts at row B + 1
    data_start = date_header_idx + 1
    if row_b_norm and col_map:
        # Check if any location was found in row B
        for c_idx in col_map:
            if c_idx < len(row_b_norm) and row_b_norm[c_idx]:
                data_start = date_header_idx + 2
                break

    # ── Step 4: Extract events ──────────────────────────────────
    events_flat = []
    current_date = None

    for row in rows[data_start:]:
        # Try to get date from date column
        if date_col < len(row):
            parsed_date = _parse_date(row[date_col])
            if parsed_date:
                current_date = parsed_date

        if current_date is None:
            continue

        # Check each location column
        for loc_col, loc_name in col_map.items():
            if loc_col >= len(row):
                continue

            event_text = _clean_event_text(row[loc_col])
            if not event_text:
                continue

            cat = ""
            sts = ""
            if category_col is not None and category_col < len(row):
                raw = row[category_col]
                if not _is_date_obj(raw):
                    cat = str(raw).strip()
                    if cat.lower() in ("nan", "none", ""):
                        cat = ""
            if status_col is not None and status_col < len(row):
                raw = row[status_col]
                if not _is_date_obj(raw):
                    sts = str(raw).strip()
                    if sts.lower() in ("nan", "none", ""):
                        sts = ""

            events_flat.append({
                "date": str(current_date),
                "event_name": event_text,
                "location": loc_name,
                "category": cat,
                "status": sts,
            })

    if len(events_flat) < 2:
        return None

    # ── Step 5: FILTER to expected month only ───────────────────
    # This prevents events from wrong months leaking in
    if expected_year and expected_month:
        expected_key = f"{expected_year}-{expected_month:02d}"
        filtered = [e for e in events_flat if e["date"].startswith(expected_key)]
        # Only apply filter if it actually matches some events
        # (protects against PERIODE parsing errors)
        if filtered:
            events_flat = filtered

    # ── Step 6: Deduplicate ─────────────────────────────────────
    seen = set()
    unique = []
    for e in events_flat:
        key = (e["date"], e["event_name"], e["location"])
        if key not in seen:
            seen.add(key)
            unique.append(e)
    unique.sort(key=lambda e: e["date"])

    # ── Step 7: Build output ────────────────────────────────────
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
    month_label = ""
    if expected_year and expected_month:
        month_label = f" ({expected_year}-{expected_month:02d})"

    return {
        "success": True,
        "format": "events",
        "is_events": True,
        "daily": daily,
        "monthly": monthly,
        "events_flat": unique,
        "message": (
            f"Events{month_label}: {len(unique)} event(s) across "
            f"{len(event_dates)} day(s)."
        ),
    }


# ── Multi-sheet reader ────────────────────────────────────────

def parse_event_file(filepath):
    """
    Read an event calendar Excel with multiple tabs (one per month).
    Tries every sheet and merges all events.
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
                    if isinstance(c, (date, datetime)):
                        converted.append(c)
                    elif hasattr(c, "date") and callable(c.date):
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
