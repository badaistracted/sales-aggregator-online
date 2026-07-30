# event_parser.py
"""
Robust parser for the mall's internal event calendar Excel file.

Guarantees:
  1. Only claims files that contain event-calendar fingerprints
     (Main Atrium, Amphitheatre, Foodtainment)
  2. Sheet-Level Month Locking: Events strictly bound to declared month
  3. Catch-All Column Mapping: Every column between Date and Category/Status
     is scanned for events
  4. Handles sheet names with or without year (e.g. "MEI" or "MEI 2026")
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

# These words MUST appear somewhere in the file for it to be an event calendar
EVENT_FINGERPRINTS = ["main atrium", "amphitheatre", "amphitheater", "foodtainment"]


def _is_date_obj(val):
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
    try:
        float(s.replace(",", "").replace(".", ""))
        return None
    except ValueError:
        pass
    return s


def _norm(val):
    if _is_date_obj(val):
        return ""
    return re.sub(r"\s+", " ", str(val).strip().lower())


def _parse_periode(text):
    """Parse 'MEI 2026' or 'PERIODE: MEI 2026'. Returns (year, month) or (None, None)."""
    if _is_date_obj(text):
        return None, None
    s = str(text).strip().lower()
    # Try "month year" pattern
    m = re.search(r"([a-z]{3,})\s+(\d{4})", s)
    if m:
        month_name = m.group(1)
        year = int(m.group(2))
        month_num = MONTH_MAP.get(month_name)
        if month_num and 2020 <= year <= 2035:
            return year, month_num
    return None, None


def _parse_month_only(text):
    """Parse just a month name like 'MEI' or 'APRIL'. Returns month number or None."""
    if _is_date_obj(text):
        return None
    s = str(text).strip().lower()
    # Must be ONLY a month name (no numbers, no other words)
    s = re.sub(r"\s+", "", s)
    return MONTH_MAP.get(s)


def _parse_date(val, expected_year=None, expected_month=None):
    """Parse date with strict anchoring to expected sheet month/year."""
    if val is None:
        return None

    dt = None

    if isinstance(val, datetime):
        dt = val.date()
    elif isinstance(val, date):
        dt = val
    elif hasattr(val, "date") and callable(val.date):
        try:
            dt = val.date()
        except Exception:
            pass
    else:
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", ""):
            return None

        # Plain day number (e.g. "1", "15") — common in merged-cell layouts
        if s.isdigit() and expected_year and expected_month:
            day = int(s)
            if 1 <= day <= 31:
                try:
                    return date(expected_year, expected_month, day)
                except ValueError:
                    pass

        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"]:
            try:
                dt = datetime.strptime(s, fmt).date()
                break
            except ValueError:
                pass

        if not dt:
            try:
                f = float(s)
                if 40000 < f < 60000:
                    base = date(1899, 12, 30)
                    dt = base + timedelta(days=int(f))
            except (ValueError, OverflowError):
                pass

    if dt and expected_year and expected_month:
        # If parsed date doesn't match expected month, try to fix it
        if dt.year == expected_year and dt.month == expected_month:
            return dt

        # Try swapping day and month (common DD/MM vs MM/DD confusion)
        if dt.day <= 12:
            try:
                swapped = date(dt.year, dt.day, dt.month)
                if swapped.year == expected_year and swapped.month == expected_month:
                    return swapped
            except ValueError:
                pass

        # Force the expected month with the parsed day
        try:
            forced = date(expected_year, expected_month, dt.day)
            return forced
        except ValueError:
            pass

        return None  # Can't reconcile — drop this date

    return dt


def _has_event_fingerprint(rows):
    """
    Check if these rows contain event calendar fingerprints.
    Returns True only if Main Atrium / Amphitheatre / Foodtainment appears.
    """
    for row in rows[:20]:
        for cell in row:
            if _is_date_obj(cell):
                continue
            n = _norm(cell)
            for fp in EVENT_FINGERPRINTS:
                if fp in n:
                    return True
    return False


def try_event_parser(rows, sheet_name="", year_hint=None):
    """
    Parse a SINGLE sheet's rows as event calendar.

    Args:
        rows: list of lists
        sheet_name: e.g. "MEI" or "MEI 2026"
        year_hint: fallback year if sheet name has no year
    """
    if not rows or len(rows) < 3:
        return None

    # Must have event fingerprint
    if not _has_event_fingerprint(rows):
        return None

    # ── Step 1: Determine expected month/year ──────────────────
    expected_year = None
    expected_month = None

    # Try sheet name with year: "MEI 2026"
    if sheet_name:
        yr, mn = _parse_periode(sheet_name)
        if yr:
            expected_year = yr
            expected_month = mn

    # Try sheet name without year: "MEI", "APRIL"
    if not expected_month and sheet_name:
        mn = _parse_month_only(sheet_name)
        if mn:
            expected_month = mn

    # Scan top rows for PERIODE header: "PERIODE: MEI 2026"
    for row in rows[:10]:
        for cell in row:
            yr, mn = _parse_periode(cell)
            if yr:
                if not expected_year:
                    expected_year = yr
                if not expected_month:
                    expected_month = mn
                break
        if expected_year and expected_month:
            break

    # If we still don't have a year, use the hint or scan for any year
    if not expected_year and year_hint:
        expected_year = year_hint

    if not expected_year:
        # Last resort: scan for any 4-digit year in top rows
        for row in rows[:15]:
            for cell in row:
                if _is_date_obj(cell):
                    continue
                m = re.search(r"(\d{4})", str(cell))
                if m:
                    yr = int(m.group(1))
                    if 2020 <= yr <= 2035:
                        expected_year = yr
                        break
            if expected_year:
                break

    if not expected_year or not expected_month:
        return None

    # ── Step 2: Find Header Row ─────────────────────────────────
    date_header_idx = None
    for idx, row in enumerate(rows[:15]):
        if any(_is_date_obj(c) for c in row):
            continue
        normed = [_norm(c) for c in row]
        for n in normed:
            if n in ("date", "tanggal", "tgl", "hari/tanggal",
                     "hari/ tanggal", "hari / tanggal") or \
               n.startswith("date") or n.startswith("tanggal"):
                date_header_idx = idx
                break
        if date_header_idx is not None:
            break

    if date_header_idx is None:
        return None

    row_a = rows[date_header_idx]
    row_a_norm = [_norm(c) for c in row_a]

    # Find Date column
    date_col = None
    for c_idx, n in enumerate(row_a_norm):
        if n in ("date", "tanggal", "tgl", "hari/tanggal",
                 "hari/ tanggal", "hari / tanggal") or \
           n.startswith("date") or n.startswith("tanggal"):
            date_col = c_idx
            break

    if date_col is None:
        return None

    # Find Category / Status columns to bound event columns
    category_col = None
    status_col = None
    for c_idx, n in enumerate(row_a_norm):
        if "categor" in n or "kategori" in n:
            category_col = c_idx
        elif "status" in n:
            status_col = c_idx

    # ── Step 3: Catch-All Column Mapping ────────────────────────
    # Skip "Time" column if it exists right after Date
    time_col = None
    for c_idx, n in enumerate(row_a_norm):
        if n in ("time", "waktu", "jam"):
            time_col = c_idx

    # Event columns = everything between Date and Category/Status,
    # excluding Time
    max_col = min(
        c for c in [category_col, status_col, len(row_a)]
        if c is not None
    )

    # Get location names from row B (sub-header row)
    row_b = rows[date_header_idx + 1] if date_header_idx + 1 < len(rows) else []
    row_b_has_locations = False
    if row_b:
        for cell in row_b:
            n = _norm(cell)
            for fp in EVENT_FINGERPRINTS:
                if fp in n:
                    row_b_has_locations = True
                    break
            if row_b_has_locations:
                break

    col_map = {}
    for c_idx in range(date_col + 1, max_col):
        if c_idx == time_col:
            continue

        loc_name = f"Area {c_idx}"

        # Try row B first (sub-headers like "Main Atrium")
        if row_b_has_locations and c_idx < len(row_b):
            b_text = _norm(row_b[c_idx])
            if b_text and b_text not in ("nan", "none", ""):
                loc_name = str(row_b[c_idx]).strip().title()

        # Fall back to row A
        elif c_idx < len(row_a):
            a_text = _norm(row_a[c_idx])
            if a_text and a_text not in ("date", "time", "tanggal", "waktu",
                                          "jam", "nan", "none", "location",
                                          "lokasi", "tempat", ""):
                loc_name = str(row_a[c_idx]).strip().title()

        col_map[c_idx] = loc_name

    if not col_map:
        return None

    # Data starts after row B if it has location sub-headers
    data_start = date_header_idx + 2 if row_b_has_locations else date_header_idx + 1

    # ── Step 4: Extract Events ──────────────────────────────────
    events_flat = []
    current_date = None

    for row in rows[data_start:]:
        # Try to parse date
        if date_col < len(row):
            parsed_date = _parse_date(row[date_col], expected_year, expected_month)
            if parsed_date:
                current_date = parsed_date

        if current_date is None:
            continue

        # STRICT: only keep events matching expected month
        if current_date.year != expected_year or current_date.month != expected_month:
            continue

        date_str = str(current_date)

        for loc_col, loc_name in col_map.items():
            if loc_col >= len(row):
                continue
            event_text = _clean_event_text(row[loc_col])
            if not event_text:
                continue

            events_flat.append({
                "date": date_str,
                "event_name": event_text,
                "location": loc_name,
                "category": "",
                "status": "",
            })

    if not events_flat:
        return None

    # Deduplicate
    seen = set()
    unique = []
    for e in events_flat:
        key = (e["date"], e["event_name"])
        if key not in seen:
            seen.add(key)
            unique.append(e)
    unique.sort(key=lambda e: e["date"])

    daily_map = defaultdict(list)
    for e in unique:
        daily_map[e["date"]].append(e["event_name"])

    daily = [
        {"date": d, "events": evts, "event_count": len(evts)}
        for d, evts in sorted(daily_map.items())
    ]

    monthly_key = f"{expected_year}-{expected_month:02d}"
    monthly = {
        monthly_key: {
            "events": unique,
            "event_count": len(unique),
            "event_days": len(set(e["date"] for e in unique)),
        }
    }

    return {
        "success": True,
        "format": "events",
        "is_events": True,
        "daily": daily,
        "monthly": monthly,
        "events_flat": unique,
        "message": f"Events ({monthly_key}): {len(unique)} event(s) across {len(daily_map)} day(s).",
    }


# ── Multi-sheet reader ────────────────────────────────────────

def parse_event_file(filepath):
    """
    Read event calendar Excel with multiple tabs.
    Only claims the file if it contains event fingerprints.
    """
    filepath = Path(filepath)

    try:
        engine = "xlrd" if filepath.suffix == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        sheet_names = xl.sheet_names
    except Exception:
        return None

    # ── Quick fingerprint check: read first sheet to verify ─────
    # This prevents claiming sales/traffic files
    try:
        df_check = pd.read_excel(xl, sheet_name=0, header=None, nrows=20).fillna("")
        check_rows = []
        for row in df_check.values.tolist():
            check_rows.append([str(c).strip() for c in row])

        if not _has_event_fingerprint(check_rows):
            # Try second sheet too (first sheet might be a cover page)
            if len(sheet_names) > 1:
                df_check2 = pd.read_excel(xl, sheet_name=1, header=None, nrows=20).fillna("")
                check_rows2 = [[str(c).strip() for c in row] for row in df_check2.values.tolist()]
                if not _has_event_fingerprint(check_rows2):
                    return None  # Not an event file
            else:
                return None  # Not an event file
    except Exception:
        return None

    # ── First pass: collect year from any sheet that has it ──────
    year_hint = None
    for sheet_name in sheet_names:
        yr, mn = _parse_periode(sheet_name)
        if yr:
            year_hint = yr
            break

    if not year_hint:
        # Scan first few sheets for PERIODE rows
        for sheet_name in sheet_names[:3]:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10).fillna("")
                for row in df.values.tolist():
                    for cell in row:
                        yr, mn = _parse_periode(cell)
                        if yr:
                            year_hint = yr
                            break
                    if year_hint:
                        break
            except Exception:
                continue
            if year_hint:
                break

    # ── Parse each sheet ────────────────────────────────────────
    all_events_flat = []
    all_monthly = {}
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

        result = try_event_parser(rows, sheet_name, year_hint=year_hint)
        if result and result.get("success"):
            all_events_flat.extend(result.get("events_flat", []))
            for mk, mv in result.get("monthly", {}).items():
                if mk not in all_monthly:
                    all_monthly[mk] = {"events": [], "event_count": 0, "event_days": 0}
                all_monthly[mk]["events"].extend(mv["events"])
                all_monthly[mk]["event_count"] += mv["event_count"]
                all_monthly[mk]["event_days"] = len(set(
                    e["date"] for e in all_monthly[mk]["events"]
                ))
            messages.append(f"Sheet '{sheet_name}': {result['message']}")

    if not all_events_flat:
        return None

    # Deduplicate
    seen = set()
    unique_flat = []
    for e in all_events_flat:
        key = (e["date"], e["event_name"])
        if key not in seen:
            seen.add(key)
            unique_flat.append(e)
    unique_flat.sort(key=lambda e: e["date"])

    daily_map = defaultdict(list)
    for e in unique_flat:
        daily_map[e["date"]].append(e["event_name"])

    daily = [
        {"date": d, "events": list(set(evts)), "event_count": len(set(evts))}
        for d, evts in sorted(daily_map.items())
    ]

    months_found = sorted(all_monthly.keys())

    return {
        "success": True,
        "format": "events",
        "is_events": True,
        "daily": daily,
        "monthly": all_monthly,
        "events_flat": unique_flat,
        "message": (
            f"Event calendar: {len(unique_flat)} event(s) across "
            f"{len(months_found)} month(s) from {len(messages)} sheet(s). "
            f"Months: {', '.join(months_found)}"
        ),
    }
