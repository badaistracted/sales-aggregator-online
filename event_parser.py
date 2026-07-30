# event_parser.py
"""
Robust parser for the mall's internal event calendar Excel file.

Guarantees:
  1. Sheet-Level Month Locking: Events are strictly bound to the month 
     declared by the sheet name or PERIODE header. No cross-month leaks.
  2. Catch-All Column Mapping: Every column between Date and Status/Category 
     is scanned for events. Zero events can be missed due to unknown headers.
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
    if _is_date_obj(text):
        return None, None
    s = str(text).strip().lower()
    m = re.search(r"([a-z]{3,})\s+(\d{4})", s)
    if m:
        month_name = m.group(1)
        year = int(m.group(2))
        month_num = MONTH_MAP.get(month_name)
        if month_num and 2020 <= year <= 2035:
            return year, month_num
    return None, None


def _parse_date(val, expected_year=None, expected_month=None):
    """
    Parse date with strict anchoring to the expected sheet month/year if provided.
    """
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

        # Check if it's just an integer day number (e.g. "1", "15") from merged cells
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

    if dt:
        # If we have an expected sheet month, ensure the parsed date matches it!
        # If day/month got flipped (e.g. 5/11 vs 11/5), correct it if possible.
        if expected_year and expected_month:
            if dt.year != expected_year or dt.month != expected_month:
                # Try swapping day and month
                try:
                    swapped = date(dt.year, dt.day, dt.month)
                    if swapped.year == expected_year and swapped.month == expected_month:
                        dt = swapped
                except ValueError:
                    pass
            # Final check: if month still doesn't match sheet, force sheet month with original day
            if dt.month != expected_month and 1 <= dt.day <= 31:
                try:
                    dt = date(expected_year, expected_month, dt.day)
                except ValueError:
                    pass
        return dt

    return None


def try_event_parser(rows, sheet_name=""):
    """
    Parse a SINGLE sheet's rows with strict sheet-level month locking
    and catch-all column mapping.
    """
    if not rows or len(rows) < 3:
        return None

    # ── Step 1: Lock to Sheet Month / Year ─────────────────────
    expected_year = None
    expected_month = None

    # Check sheet name first (e.g. "MEI 2026")
    if sheet_name:
        yr, mn = _parse_periode(sheet_name)
        if yr:
            expected_year = yr
            expected_month = mn

    # If sheet name didn't give a month, scan top rows for PERIODE
    if not expected_year:
        for row in rows[:10]:
            for cell in row:
                yr, mn = _parse_periode(cell)
                if yr:
                    expected_year = yr
                    expected_month = mn
                    break
            if expected_year:
                break

    if not expected_year or not expected_month:
        return None  # Can't determine which month this sheet belongs to

    # ── Step 2: Find Header Row ─────────────────────────────────
    date_header_idx = None
    for idx, row in enumerate(rows[:15]):
        if any(_is_date_obj(c) for c in row):
            continue
        normed = [_norm(c) for c in row]
        if any(n in ("date", "tanggal", "tgl", "hari/tanggal") or n.startswith("date") for n in normed):
            date_header_idx = idx
            break

    if date_header_idx is None:
        return None

    row_a = rows[date_header_idx]
    row_a_norm = [_norm(c) for c in row_a]

    # Find Date column
    date_col = None
    for c_idx, n in enumerate(row_a_norm):
        if n in ("date", "tanggal", "tgl", "hari/tanggal") or n.startswith("date"):
            date_col = c_idx
            break

    if date_col is None:
        return None

    # Find Category / Status columns to bound our event columns
    category_col = None
    status_col = None
    for c_idx, n in enumerate(row_a_norm):
        if "categor" in n or "kategori" in n:
            category_col = c_idx
        elif "status" in n:
            status_col = c_idx

    # ── Step 3: Catch-All Column Mapping ────────────────────────
    # EVERY column between Date and Category/Status is an event column.
    max_col = min(
        category_col if category_col is not None else 999,
        status_col if status_col is not None else 999,
        len(row_a)
    )

    col_map = {}
    row_b = rows[date_header_idx + 1] if date_header_idx + 1 < len(rows) else []

    for c_idx in range(date_col + 1, max_col):
        # Try to get a location name from Row B, then Row A, else default
        loc_name = f"Area {c_idx}"
        if c_idx < len(row_b) and str(row_b[c_idx]).strip():
            loc_name = str(row_b[c_idx]).strip().title()
        elif c_idx < len(row_a) and str(row_a[c_idx]).strip():
            txt = str(row_a[c_idx]).strip()
            if txt.lower() not in ("date", "time", "tanggal", "waktu"):
                loc_name = txt.title()
        col_map[c_idx] = loc_name

    if not col_map:
        return None

    data_start = date_header_idx + 2 if row_b else date_header_idx + 1

    # ── Step 4: Extract Events with Strict Month Enforcement ────
    events_flat = []
    current_date = None

    for row in rows[data_start:]:
        if date_col < len(row):
            parsed_date = _parse_date(row[date_col], expected_year, expected_month)
            if parsed_date:
                current_date = parsed_date

        if current_date is None:
            continue

        # STRICT ENFORCEMENT: Drop any row whose date doesn't match the sheet's month
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
        {
            "date": d,
            "events": evts,
            "event_count": len(evts),
        }
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
        "message": f"Events ({expected_year}-{expected_month:02d}): {len(unique)} event(s) extracted.",
    }


# ── Multi-sheet reader ────────────────────────────────────────

def parse_event_file(filepath):
    filepath = Path(filepath)

    try:
        engine = "xlrd" if filepath.suffix == ".xls" else "openpyxl"
        xl = pd.ExcelFile(filepath, engine=engine)
        sheet_names = xl.sheet_names
    except Exception:
        return None

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

        result = try_event_parser(rows, sheet_name)
        if result and result.get("success"):
            all_events_flat.extend(result.get("events_flat", []))
            for mk, mv in result.get("monthly", {}).items():
                if mk not in all_monthly:
                    all_monthly[mk] = {"events": [], "event_count": 0, "event_days": 0}
                all_monthly[mk]["events"].extend(mv["events"])
                all_monthly[mk]["event_count"] += mv["event_count"]
                all_monthly[mk]["event_days"] = len(set(e["date"] for e in all_monthly[mk]["events"]))
            messages.append(f"Sheet '{sheet_name}': {result['message']}")

    if not all_events_flat:
        return None

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
        {"date": d, "events": evts, "event_count": len(evts)}
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
        "message": f"Event calendar: {len(unique_flat)} event(s) across {len(months_found)} month(s). Sheets: {', '.join(sheet_names)}",
    }
