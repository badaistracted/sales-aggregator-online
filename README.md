# Tenant Sales Aggregator

A local web app that aggregates daily sales from multiple tenant Excel files,
applies Indonesian calendar logic (weekends + national holidays), and produces
a formatted multi-tab Excel report plus an in-browser dashboard.

## Setup

```bash
cd tenant_sales_aggregator
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:5000** in your browser.

No internet connection is required once the packages above are installed —
Chart.js is vendored locally in `static/chart.umd.js` and the UI uses system
fonts, so it works fully offline (useful if you're running this on a mall
office machine with restricted network access).

## How to use it

1. Drag and drop (or browse for) one or more tenant `.xlsx`/`.xls` files.
   Each file needs at least a **Date** column and a **Sales** column (a few
   common variants like "Tanggal" / "Revenue" are also recognized).
2. Click **Process files**. The dashboard shows the Monthly_Summary table
   and a chart per tenant, with weekend/holiday days marked in green.
3. Click **Export integrated report (.xlsx)** to download the full workbook.

Files that are missing a required column, or have no valid dates, are
skipped with an explanation shown on screen — the rest of the batch still
processes normally.

## What's inside

| File | Purpose |
|---|---|
| `backend/data_parser.py` | Phase 1 — reads tenant files, builds a continuous daily timeline for the month (missing dates filled with 0 sales), classifies each day as Weekday/Weekend using the Indonesian public holiday calendar (`holidays` library, `country='ID'`) |
| `backend/excel_export.py` | Phases 2–3 — builds the 6-sheet workbook (`Monthly_Summary`, one sheet per tenant, `All_Tenants_Weekends`, `All_Tenants_Weekdays`) with live formulas and a native Excel line chart per tenant sheet |
| `app.py` | Phase 4 — Flask server: upload endpoint, processing pipeline, download endpoint |
| `templates/index.html`, `static/` | Phase 5 — the upload zone + dashboard UI (summary table, per-tenant Chart.js charts, error banners) |

## Assumptions worth knowing about

- **Missing dates are filled with 0 sales**, not interpolated or dropped —
  the spec asked for a continuous timeline with no gaps, and 0 was the
  least presumptive way to represent an unreported/closed day. If your
  tenants sometimes skip reporting rather than actually having zero sales,
  you may want to review those flat days before trusting the averages.
- **A date that appears twice in one tenant's file is summed**, not
  overwritten, on the assumption that both rows are same-day sales entries.
- **`Monthly_Summary` uses live Excel formulas** (`SUM`, `AVERAGEIFS`)
  referencing each tenant's own sheet, so the numbers recalculate
  automatically if you edit a tenant sheet after export.
- The `All_Tenants_Weekends` / `All_Tenants_Weekdays` sheets are filtered
  snapshots (not formulas) — regenerate the report if the underlying data
  changes.

## Sample data

`sample_data/` has a few generated example tenant files (including one
intentionally missing its Sales column) if you want to try the app before
pointing it at real K Square exports.
