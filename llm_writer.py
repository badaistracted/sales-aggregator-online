# llm_writer.py
"""
LLM-powered slide text generation.
Python calculates all numbers — the LLM only writes commentary.

Falls back to template text if no API key is set.
"""
import os
import json
import calendar as cal


def _fmt(v: float) -> str:
    """Format number as IDR string for LLM prompt."""
    if v >= 1_000_000_000:
        return f"IDR {v/1_000_000_000:.2f} billion"
    if v >= 1_000_000:
        return f"IDR {v/1_000_000:.1f} million"
    return f"IDR {v:,.0f}"


def _build_kpis(master_data: dict, traffic_data: dict | None,
                target_year: int, target_month: int,
                events_data: dict | None = None) -> dict:
    """
    Compute all KPIs from parsed data.
    This is pure Python — no LLM involved here.

    Returns a flat dict of named metrics.
    """
    target_key  = f"{target_year}-{target_month:02d}"
    month_label = f"{cal.month_name[target_month]} {target_year}"

    # Build previous month key
    if target_month == 1:
        prev_key = f"{target_year - 1}-12"
    else:
        prev_key = f"{target_year}-{target_month - 1:02d}"

    # ── Sales KPIs ──────────────────────────────────────────
    tenant_sales = {}
    tenant_prev  = {}
    all_months   = set()

    for tenant, tm in master_data.items():
        monthly = tm.get("monthly", {})
        all_months.update(monthly.keys())
        tenant_sales[tenant] = monthly.get(target_key, 0)
        tenant_prev[tenant]  = monthly.get(prev_key, 0)

    total_sales      = sum(tenant_sales.values())
    total_prev_sales = sum(tenant_prev.values())
    sales_growth     = (
        (total_sales - total_prev_sales) / total_prev_sales * 100
        if total_prev_sales > 0 else None
    )

    sorted_tenants = sorted(tenant_sales.items(), key=lambda x: x[1], reverse=True)
    top_tenant     = sorted_tenants[0]  if sorted_tenants else ("N/A", 0)
    bottom_tenant  = sorted_tenants[-1] if sorted_tenants else ("N/A", 0)

    # Tenants that improved vs previous month
    improvers = [
        t for t in master_data
        if tenant_sales.get(t, 0) > tenant_prev.get(t, 0)
        and tenant_prev.get(t, 0) > 0
    ]
    decliners = [
        t for t in master_data
        if tenant_sales.get(t, 0) < tenant_prev.get(t, 0)
        and tenant_prev.get(t, 0) > 0
    ]

    # ── Traffic KPIs ─────────────────────────────────────────
    total_traffic   = 0
    prev_traffic    = 0
    traffic_growth  = None
    sales_per_visitor = None

    if traffic_data and traffic_data.get("monthly"):
        total_traffic  = traffic_data["monthly"].get(target_key, 0)
        prev_traffic   = traffic_data["monthly"].get(prev_key, 0)
        if prev_traffic > 0:
            traffic_growth = (total_traffic - prev_traffic) / prev_traffic * 100
        if total_traffic > 0:
            sales_per_visitor = total_sales / total_traffic

    # ── Daily KPIs ────────────────────────────────────────────
    from collections import defaultdict
    from datetime import datetime

    all_daily = []
    for tm in master_data.values():
        all_daily.extend(tm.get("daily", []))

    target_daily = [d for d in all_daily if d["date"].startswith(target_key)]

    by_date = defaultdict(float)
    for d in target_daily:
        by_date[d["date"]] += d["sales"]

    peak_day   = max(by_date.items(), key=lambda x: x[1]) if by_date else (None, 0)
    lowest_day = min(by_date.items(), key=lambda x: x[1]) if by_date else (None, 0)

    weekday_sales  = []
    weekend_sales  = []
    for date_str, val in by_date.items():
        try:
            dow = datetime.strptime(date_str, "%Y-%m-%d").weekday()
        except ValueError:
            continue
        if dow >= 5:
            weekend_sales.append(val)
        else:
            weekday_sales.append(val)

    avg_weekday = sum(weekday_sales) / len(weekday_sales) if weekday_sales else 0
    avg_weekend = sum(weekend_sales) / len(weekend_sales) if weekend_sales else 0

    # ── Event Impact KPIs ─────────────────────────────────────
    event_impact = []
    if events_data and events_data.get("daily"):
        # Build event lookup: date → list of event names
        event_dates = {}
        for ed in events_data["daily"]:
            event_dates[ed["date"]] = ed.get("events", [])

        # Split sales into event days vs non-event days
        event_day_sales = []
        non_event_day_sales = []

        for date_str, sales_val in by_date.items():
            if date_str in event_dates and event_dates[date_str]:
                event_day_sales.append({
                    "date": date_str,
                    "sales": sales_val,
                    "events": event_dates[date_str],
                })
            else:
                non_event_day_sales.append(sales_val)

        avg_non_event = (
            sum(non_event_day_sales) / len(non_event_day_sales)
            if non_event_day_sales else 0
        )
        avg_event = (
            sum(d["sales"] for d in event_day_sales) / len(event_day_sales)
            if event_day_sales else 0
        )

        # Rank event days by sales uplift vs non-event average
        for ed in event_day_sales:
            uplift = ed["sales"] - avg_non_event
            uplift_pct = (
                (uplift / avg_non_event * 100) if avg_non_event > 0 else 0
            )
            event_impact.append({
                "date": ed["date"],
                "sales": ed["sales"],
                "events": ed["events"],
                "uplift": uplift,
                "uplift_pct": uplift_pct,
            })

        event_impact.sort(key=lambda x: x["sales"], reverse=True)
                    
    return {
        # Identity
        "month_label"         : month_label,
        "target_key"          : target_key,
        "prev_key"            : prev_key,
        "total_tenants"       : len(master_data),

        # Sales
        "total_sales"         : total_sales,
        "total_prev_sales"    : total_prev_sales,
        "sales_growth_pct"    : sales_growth,

        # Tenants
        "top_tenant_name"     : top_tenant[0],
        "top_tenant_sales"    : top_tenant[1],
        "bottom_tenant_name"  : bottom_tenant[0],
        "bottom_tenant_sales" : bottom_tenant[1],
        "tenant_sales"        : dict(sorted_tenants),
        "improvers"           : improvers,
        "decliners"           : decliners,

        # Traffic
        "total_traffic"       : total_traffic,
        "prev_traffic"        : prev_traffic,
        "traffic_growth_pct"  : traffic_growth,
        "sales_per_visitor"   : sales_per_visitor,

        # Daily
        "peak_day"            : peak_day[0],
        "peak_sales"          : peak_day[1],
        "lowest_day"          : lowest_day[0],
        "lowest_sales"        : lowest_day[1],
        "avg_weekday_sales"   : avg_weekday,
        "avg_weekend_sales"   : avg_weekend,
        "daily_days_tracked"  : len(by_date),

        # Initialize event defaults
        event_impact = []
        avg_event = 0
        avg_non_event = 0
        event_day_sales = []
        non_event_day_sales = []

        # Event impact
        "event_impact"        : event_impact,
        "avg_event_day_sales" : avg_event if event_impact else 0,
        "avg_non_event_sales" : avg_non_event if event_impact else 0,
        "event_days_count"    : len(event_day_sales) if event_impact else 0,
        "non_event_days_count": len(non_event_day_sales) if event_impact else 0,
    }


def _template_summary(kpis: dict) -> dict:
    """
    Fallback template text — used when no OpenAI key is configured.
    Returns same structure as LLM response.
    """
    k  = kpis
    ml = k["month_label"]

    growth_str = (
        f"{k['sales_growth_pct']:+.1f}% vs prior month"
        if k["sales_growth_pct"] is not None
        else "prior month comparison not available"
    )

    traffic_str = (
        f"Visitor count: {k['total_traffic']:,.0f} ({k['traffic_growth_pct']:+.1f}% MoM). "
        f"Sales per visitor: {_fmt(k['sales_per_visitor'])}."
        if k["total_traffic"] > 0 and k["sales_per_visitor"]
        else "Traffic data not available for this period."
    )

    peak_str = (
        f"Peak day was {k['peak_day']} with {_fmt(k['peak_sales'])} in sales."
        if k["peak_day"] else ""
    )

    return {
        "executive_summary": (
            f"In {ml}, total tenant sales reached {_fmt(k['total_sales'])} "
            f"({growth_str}). "
            f"{k['total_tenants']} tenants contributed to the result. "
            f"Top performer: {k['top_tenant_name']} at {_fmt(k['top_tenant_sales'])}. "
            f"{traffic_str} {peak_str}"
        ).strip(),

        "sales_slide_notes": (
            f"Total sales of {_fmt(k['total_sales'])} recorded in {ml}. "
            f"{'Sales improved month-over-month.' if (k['sales_growth_pct'] or 0) > 0 else 'Sales declined versus prior month.'}"
        ),

        "tenant_slide_notes": (
            f"{k['top_tenant_name']} leads with {_fmt(k['top_tenant_sales'])}. "
            f"{len(k['improvers'])} tenant(s) improved vs last month; "
            f"{len(k['decliners'])} declined."
        ),

        "traffic_slide_notes": (
            traffic_str
        ),

        "daily_slide_notes": (
            f"Weekend average: {_fmt(k['avg_weekend_sales'])} vs "
            f"weekday average: {_fmt(k['avg_weekday_sales'])}. "
            f"{peak_str}"
        ) if k["avg_weekend_sales"] > 0 else "Daily breakdown available in appendix.",

        "recommendations": [
            f"Focus support on {k['bottom_tenant_name']} — lowest sales this month.",
            "Increase weekend promotions to leverage higher foot traffic." if k["avg_weekend_sales"] > k["avg_weekday_sales"] else "Investigate weekday traffic decline and consider targeted promotions.",
            f"Replicate success factors from {k['top_tenant_name']} across lower-performing tenants.",
            "Review tenant mix and lease terms for consistently underperforming units.",
        ],
    }


def generate_slide_text(kpis: dict) -> dict:
    """
    Main entry point.
    Uses OpenAI if OPENAI_API_KEY is set, otherwise uses template fallback.

    Always returns:
    {
        "executive_summary": str,
        "sales_slide_notes": str,
        "tenant_slide_notes": str,
        "traffic_slide_notes": str,
        "daily_slide_notes": str,
        "recommendations": list[str],
    }
    """
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()

    if not api_key:
        print("[llm_writer] No OPENAI_API_KEY — using template fallback.")
        return _template_summary(kpis)

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        # Build a compact context for the LLM
        # We pass numbers already formatted — LLM does NOT calculate
        growth = (
            f"{kpis['sales_growth_pct']:+.1f}%"
            if kpis["sales_growth_pct"] is not None else "N/A"
        )
        traffic_ctx = ""
        if kpis["total_traffic"] > 0:
            traffic_ctx = (
                f"- Visitor count: {kpis['total_traffic']:,.0f} "
                f"({kpis['traffic_growth_pct']:+.1f}% MoM)\n"
                f"- Sales per visitor: {_fmt(kpis['sales_per_visitor'])}\n"
            )

        top5 = list(kpis["tenant_sales"].items())[:5]
        top5_str = "\n".join(
            f"  {i+1}. {name}: {_fmt(val)}"
            for i, (name, val) in enumerate(top5)
        )

        prompt = f"""You are writing a monthly performance report for a shopping mall management team.
Be concise, professional, and business-focused. Use specific numbers provided.
Do NOT invent numbers. Do NOT use markdown. Write in complete sentences.

REPORT MONTH: {kpis['month_label']}
TOTAL SALES: {_fmt(kpis['total_sales'])} ({growth} vs prior month)
TOTAL TENANTS: {kpis['total_tenants']}
{traffic_ctx}
TOP 5 TENANTS:
{top5_str}
IMPROVERS vs last month: {len(kpis['improvers'])} tenants
DECLINERS vs last month: {len(kpis['decliners'])} tenants
PEAK DAY: {kpis['peak_day']} — {_fmt(kpis['peak_sales'])}
AVG WEEKDAY SALES: {_fmt(kpis['avg_weekday_sales'])}
AVG WEEKEND SALES: {_fmt(kpis['avg_weekend_sales'])}

Write a JSON object with these exact keys:
{{
  "executive_summary": "3-4 sentence overview for senior management",
  "sales_slide_notes": "2 sentences about monthly sales trend",
  "tenant_slide_notes": "2 sentences about tenant performance",
  "traffic_slide_notes": "2 sentences about visitor traffic and spend",
  "daily_slide_notes": "2 sentences about daily sales patterns",
  "recommendations": ["action item 1", "action item 2", "action item 3", "action item 4"]
}}

Return ONLY the JSON object. No extra text."""

        response = client.chat.completions.create(
            model="gpt-4o-mini",          # Fast and cheap for structured output
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4,              # Low temperature = consistent, factual
            max_tokens=800,
            response_format={"type": "json_object"},
        )

        raw = response.choices[0].message.content
        parsed = json.loads(raw)

        # Validate keys — fall back to template for any missing field
        required = [
            "executive_summary", "sales_slide_notes", "tenant_slide_notes",
            "traffic_slide_notes", "daily_slide_notes", "recommendations",
        ]
        for key in required:
            if key not in parsed:
                print(f"[llm_writer] Missing key '{key}' — using template.")
                return _template_summary(kpis)

        return parsed

    except Exception as e:
        print(f"[llm_writer] LLM call failed: {e} — using template fallback.")
        return _template_summary(kpis)
