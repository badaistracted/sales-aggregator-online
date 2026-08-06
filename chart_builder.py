# chart_builder.py
"""
Pure Python chart generation using matplotlib.
All charts return BytesIO objects — no files written to disk.
Colors match the dark-blue brand palette of the UI.
"""
import io
import calendar as cal
import matplotlib
matplotlib.use("Agg")  # Non-interactive backend — required for server use
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

# ── Brand palette ──────────────────────────────────────────────
BLUE_DARK   = "#1A3A5C"
BLUE_MID    = "#2874A6"
BLUE_LIGHT  = "#D6EAF8"
ACCENT_TEAL = "#1ABC9C"
ACCENT_AMB  = "#F39C12"
ACCENT_RED  = "#E74C3C"
ACCENT_PCT  = "#685BC7"   # purple — used for MoM % sticker (matches UI accent)
TEXT_DARK   = "#2C3E50"
GRID_COLOR  = "#E5EAF0"
BG_WHITE    = "#FFFFFF"

MONTH_SHORT = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# 18-colour palette for the all-tenants line chart
_TENANT_COLORS = [
    "#2874A6", "#1ABC9C", "#E74C3C", "#F39C12", "#8E44AD",
    "#16A085", "#2980B9", "#D35400", "#27AE60", "#C0392B",
    "#2C3E50", "#F1C40F", "#7F8C8D", "#6C3483", "#117A65",
    "#1A5276", "#A04000", "#1E8449",
]


def _base_fig(w=10, h=5):
    """Standard figure with white background."""
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor(BG_WHITE)
    ax.set_facecolor(BG_WHITE)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)
    ax.tick_params(colors=TEXT_DARK, labelsize=9)
    return fig, ax


def _save(fig):
    """Save figure to BytesIO and close."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=BG_WHITE)
    plt.close(fig)
    buf.seek(0)
    return buf


def _fmt_idr(v, _pos=None):
    """Compact IDR formatter: 1.2M, 500K, etc."""
    if v >= 1_000_000_000:
        return f"{v/1_000_000_000:.1f}B"
    if v >= 1_000_000:
        return f"{v/1_000_000:.1f}M"
    if v >= 1_000:
        return f"{v/1_000:.0f}K"
    return str(int(v))


def _month_label(key):
    """'2026-07' → 'Jul-26'"""
    parts = key.split("-")
    return f"{MONTH_SHORT[int(parts[1])]}-{parts[0][2:]}"


# ── Chart 1: Monthly Sales Bar Chart ──────────────────────────

def chart_monthly_sales(monthly_totals: dict, target_key: str) -> io.BytesIO:
    """
    Bar chart: total sales per month across all tenants.

    monthly_totals = {"2026-01": 5_000_000, "2026-02": 6_200_000, ...}
    target_key     = "2026-05"
    """
    months = sorted(monthly_totals.keys())
    values = [monthly_totals[m] for m in months]
    labels = [_month_label(m) for m in months]

    fig, ax = _base_fig(10, 5)

    colors = [BLUE_MID if m != target_key else ACCENT_TEAL for m in months]
    bars = ax.bar(labels, values, color=colors, width=0.6,
                  zorder=3, edgecolor="white", linewidth=0.5)

    # Value labels on top of each bar
    for bar, val in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + max(values) * 0.01,
            _fmt_idr(val),
            ha="center", va="bottom",
            fontsize=8, color=TEXT_DARK, fontweight="bold",
        )

    ax.set_title("Total Sales by Month", fontsize=13, fontweight="bold",
                 color=BLUE_DARK, pad=15)
    ax.set_ylabel("Sales (IDR)", fontsize=9, color=TEXT_DARK)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax.set_ylim(0, max(values) * 1.18)

    from matplotlib.patches import Patch
    legend = [
        Patch(color=BLUE_MID, label="Other months"),
        Patch(color=ACCENT_TEAL, label="Target month"),
    ]
    ax.legend(handles=legend, fontsize=8, framealpha=0.7, loc="upper left")

    fig.tight_layout()
    return _save(fig)


# ── Chart 2: Top Tenants Horizontal Bar ───────────────────────

def chart_top_tenants(
    tenant_sales: dict,
    target_key: str,
    top_n: int = 10,
) -> io.BytesIO:
    """
    Horizontal bar: top N tenants by sales in target month.
    """
    sorted_items = sorted(tenant_sales.items(), key=lambda x: x[1], reverse=True)
    items = sorted_items[:top_n]
    if not items:
        return None

    names  = [x[0] for x in reversed(items)]
    values = [x[1] for x in reversed(items)]

    n = len(items)
    colors = []
    for i in range(n):
        if i == n - 1:
            colors.append(ACCENT_TEAL)
        elif i == n - 2:
            colors.append(BLUE_MID)
        else:
            ratio = i / max(n - 2, 1)
            colors.append(
                f"#{int(0x2E + ratio * (0x93 - 0x2E)):02X}"
                f"{int(0x86 + ratio * (0xC5 - 0x86)):02X}"
                f"{int(0xC1 + ratio * (0xFD - 0xC1)):02X}"
            )

    h = max(4, n * 0.55)
    fig, ax = plt.subplots(figsize=(10, h))
    fig.patch.set_facecolor(BG_WHITE)
    ax.set_facecolor(BG_WHITE)

    bars = ax.barh(names, values, color=colors, height=0.65,
                   edgecolor="white", linewidth=0.5, zorder=3)

    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.8, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)

    max_val = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(
            val + max_val * 0.01,
            bar.get_y() + bar.get_height() / 2,
            _fmt_idr(val),
            va="center", ha="left",
            fontsize=8, color=TEXT_DARK, fontweight="bold",
        )

    ax.set_xlim(0, max_val * 1.20)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax.tick_params(axis="y", labelsize=9, colors=TEXT_DARK)
    ax.tick_params(axis="x", labelsize=8, colors=TEXT_DARK)

    target_label = _month_label(target_key)
    ax.set_title(f"Top {n} Tenants — {target_label}",
                 fontsize=13, fontweight="bold", color=BLUE_DARK, pad=12)
    ax.set_xlabel("Sales (IDR)", fontsize=9, color=TEXT_DARK)

    fig.tight_layout()
    return _save(fig)


# ── Chart 3: Traffic + Sales/Visitor Line Chart ───────────────

def chart_traffic(
    traffic_monthly: dict,
    sales_monthly: dict,
    target_key: str,
) -> io.BytesIO:
    """
    Dual-axis chart: visitor count (bar) + sales per visitor (line).
    """
    months = sorted(set(list(traffic_monthly.keys()) + list(sales_monthly.keys())))
    if not months:
        return None

    labels = [_month_label(m) for m in months]
    traffic_vals = [traffic_monthly.get(m, 0) for m in months]
    spv_vals = []
    for m in months:
        t = traffic_monthly.get(m, 0)
        s = sales_monthly.get(m, 0)
        spv_vals.append(round(s / t) if t > 0 else 0)

    fig, ax1 = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor(BG_WHITE)
    ax1.set_facecolor(BG_WHITE)

    x = np.arange(len(months))

    bar_colors = [BLUE_LIGHT if m != target_key else BLUE_MID for m in months]
    ax1.bar(x, traffic_vals, color=bar_colors, width=0.55,
            zorder=3, edgecolor="white", linewidth=0.5, label="Visitors")

    ax1.set_ylabel("Visitor Count", fontsize=9, color=BLUE_DARK)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax1.tick_params(axis="y", colors=BLUE_DARK)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)

    ax2 = ax1.twinx()
    ax2.plot(x, spv_vals, color=ACCENT_AMB, linewidth=2.5,
             marker="o", markersize=6, zorder=5, label="Sales/Visitor")
    ax2.set_ylabel("Sales per Visitor (IDR)", fontsize=9, color=ACCENT_AMB)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax2.tick_params(axis="y", colors=ACCENT_AMB)
    ax2.spines["top"].set_visible(False)

    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, fontsize=9, color=TEXT_DARK)
    ax1.set_title("Traffic & Sales per Visitor", fontsize=13,
                  fontweight="bold", color=BLUE_DARK, pad=12)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2,
               fontsize=8, loc="upper left", framealpha=0.7)

    fig.tight_layout()
    return _save(fig)


# ── Chart 4: Daily Sales Line Chart ───────────────────────────

def chart_daily_sales(
    daily_data: list,
    target_key: str,
    traffic_daily: list = None,
    tenant_name: str = "All Tenants",
) -> io.BytesIO:
    """
    Line chart of daily sales for the target month.
    """
    filtered = [d for d in daily_data if d["date"].startswith(target_key)]
    if not filtered:
        return None

    from collections import defaultdict
    by_date = defaultdict(float)
    for d in filtered:
        by_date[d["date"]] += d["sales"]

    dates  = sorted(by_date.keys())
    values = [by_date[d] for d in dates]
    x_labels = [d[8:10] for d in dates]

    traffic_vals = None
    if traffic_daily:
        traffic_by_date = defaultdict(float)
        for d in traffic_daily:
            if d["date"].startswith(target_key):
                traffic_by_date[d["date"]] += d["traffic"]
        if traffic_by_date:
            traffic_vals = [traffic_by_date.get(d, 0) for d in dates]

    from datetime import datetime as dt_cls
    weekend_mask = [dt_cls.strptime(d, "%Y-%m-%d").weekday() >= 5 for d in dates]

    fig, ax1 = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor(BG_WHITE)
    ax1.set_facecolor(BG_WHITE)

    x = np.arange(len(dates))

    for i, is_we in enumerate(weekend_mask):
        if is_we:
            ax1.axvspan(i - 0.5, i + 0.5, color="#E8F5E9", zorder=0, alpha=0.7)

    ax1.plot(x, values, color=BLUE_MID, linewidth=2.2, zorder=4, label="Sales")
    ax1.fill_between(x, values, alpha=0.12, color=BLUE_MID, zorder=3)

    for i, (val, is_we) in enumerate(zip(values, weekend_mask)):
        color = ACCENT_TEAL if is_we else BLUE_MID
        ax1.plot(i, val, "o", color=color, markersize=5, zorder=5)

    ax1.set_ylabel("Sales (IDR)", fontsize=9, color=BLUE_MID)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax1.tick_params(axis="y", colors=BLUE_MID)
    ax1.spines["top"].set_visible(False)
    ax1.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)

    if traffic_vals:
        ax2 = ax1.twinx()
        ax2.plot(x, traffic_vals, color=ACCENT_AMB, linewidth=2,
                 marker="s", markersize=4, zorder=5, alpha=0.85, label="Traffic")
        ax2.set_ylabel("Visitor Count", fontsize=9, color=ACCENT_AMB)
        ax2.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
        ax2.tick_params(axis="y", colors=ACCENT_AMB)
        ax2.spines["top"].set_visible(False)
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2,
                   fontsize=8, loc="upper left", framealpha=0.7)
    else:
        ax1.legend(fontsize=8, loc="upper left", framealpha=0.7)
        ax1.spines["right"].set_visible(False)

    ax1.spines["left"].set_color(GRID_COLOR)
    ax1.spines["bottom"].set_color(GRID_COLOR)
    ax1.set_xticks(x)
    ax1.set_xticklabels(x_labels, fontsize=7.5, color=TEXT_DARK)

    m_num = int(target_key.split("-")[1])
    y_num = target_key.split("-")[0]
    ax1.set_title(
        f"Daily Sales & Traffic — {cal.month_name[m_num]} {y_num}  ({tenant_name})",
        fontsize=13, fontweight="bold", color=BLUE_DARK, pad=12,
    )
    ax1.set_xlabel("Day of Month", fontsize=9, color=TEXT_DARK)

    if values:
        peak_i = int(np.argmax(values))
        ax1.annotate(
            f"Peak\n{_fmt_idr(values[peak_i])}",
            xy=(peak_i, values[peak_i]),
            xytext=(peak_i, values[peak_i] + max(values) * 0.08),
            ha="center", fontsize=7.5, color=ACCENT_AMB,
            arrowprops=dict(arrowstyle="-", color=ACCENT_AMB, lw=1),
        )

    fig.tight_layout()
    return _save(fig)


# ══════════════════════════════════════════════════════════════
#  Chart 5: All-Tenants Combined Line Chart
#  (for percentage_summary data)
# ══════════════════════════════════════════════════════════════

def chart_all_tenants_line(
    tenants_data: dict,
    all_months: list,
    pct_data: dict,
    target_key: str,
) -> io.BytesIO:
    """
    One line per tenant, all plotted on the same axes.
    Last data point on each line gets a circle annotation
    showing MoM % change (colour: #685BC7, no red/green coding per Q20).

    tenants_data = { "GOGO": {"monthly": {"2026-06": ..., ...}, ...}, ... }
    all_months   = ["2025-12", "2026-01", ..., "2026-07"]  (sorted)
    pct_data     = { "GOGO": {"from": "2026-06", "to": "2026-07", "pct": 5.0}, ... }
    target_key   = "2026-07"
    """
    if not tenants_data or not all_months:
        return None

    x = np.arange(len(all_months))
    x_labels = [_month_label(m) for m in all_months]

    # Wider figure to fit legend below + many months on x-axis
    fig, ax = plt.subplots(figsize=(13, 6))
    fig.patch.set_facecolor(BG_WHITE)
    ax.set_facecolor(BG_WHITE)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)

    tenant_names = list(tenants_data.keys())

    for i, tenant in enumerate(tenant_names):
        monthly  = tenants_data[tenant].get("monthly", {})
        color    = _TENANT_COLORS[i % len(_TENANT_COLORS)]
        y_plot   = [
            monthly[m] if m in monthly else np.nan
            for m in all_months
        ]

        ax.plot(x, y_plot, color=color, linewidth=1.6,
                marker="o", markersize=3, alpha=0.85,
                label=tenant, zorder=2)

        # ── MoM % sticker on the last valid data point ────────
        last_idx = next(
            (j for j in range(len(y_plot) - 1, -1, -1)
             if not np.isnan(y_plot[j])),
            None,
        )
        if last_idx is None:
            continue

        last_val = y_plot[last_idx]

        # Get pct: from pct_data first, then compute from previous point
        pct_val = pct_data.get(tenant, {}).get("pct") if pct_data else None
        if pct_val is None and last_idx > 0:
            prev = y_plot[last_idx - 1]
            if not np.isnan(prev) and prev > 0:
                pct_val = round((last_val - prev) / prev * 100, 1)

        # Small filled circle in brand colour
        ax.plot(last_idx, last_val, "o",
                color=ACCENT_PCT, markersize=11,
                markeredgecolor="white", markeredgewidth=1.2,
                zorder=5)

        if pct_val is not None:
            sign = "+" if pct_val > 0 else ""
            ax.annotate(
                f"{sign}{pct_val:.0f}%",
                xy=(last_idx, last_val),
                xytext=(0, 0),
                textcoords="offset points",
                fontsize=5.5, fontweight="bold",
                color="white",
                ha="center", va="center",
                zorder=6,
            )

    # X-axis — rotate if many months
    rotation = 45 if len(all_months) > 6 else 0
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, fontsize=8, color=TEXT_DARK,
                       rotation=rotation, ha="right" if rotation else "center")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax.tick_params(axis="y", colors=TEXT_DARK, labelsize=8)

    ax.set_title("All Tenants — Monthly Sales Trend",
                 fontsize=13, fontweight="bold", color=BLUE_DARK, pad=14)
    ax.set_ylabel("Sales (IDR)", fontsize=9, color=TEXT_DARK)

    # Legend below the chart, 4 columns
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=4, fontsize=7,
        frameon=True, framealpha=0.6,
        labelcolor=TEXT_DARK,
    )

    fig.tight_layout()
    return _save(fig)


# ══════════════════════════════════════════════════════════════
#  Chart 6: Single-Tenant Mini Line Chart
#  (for 6-per-slide grid in percentage_summary slides)
# ══════════════════════════════════════════════════════════════

def chart_single_tenant_line(
    tenant_name: str,
    monthly_data: dict,
    all_months: list,
    pct_val: float = None,
    target_key: str = None,
) -> io.BytesIO:
    """
    Compact line chart for one tenant.
    Designed to be placed in a 3×2 grid on a slide (small figsize).

    pct_val   — pre-computed MoM % (or None to auto-compute from data)
    target_key — highlighted month key, e.g. "2026-07"
    """
    if not monthly_data or not all_months:
        return None

    x        = np.arange(len(all_months))
    x_labels = [_month_label(m) for m in all_months]
    y_plot   = [
        monthly_data[m] if m in monthly_data else np.nan
        for m in all_months
    ]

    # Small figure — will be embedded 3×2 per slide
    fig, ax = plt.subplots(figsize=(4.0, 2.5))
    fig.patch.set_facecolor(BG_WHITE)
    ax.set_facecolor(BG_WHITE)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.5, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)

    ax.plot(x, y_plot, color=BLUE_MID, linewidth=1.8,
            marker="o", markersize=3.5, zorder=3)
    ax.fill_between(x, y_plot, alpha=0.08, color=BLUE_MID, zorder=2)

    # Highlight target month bar (vertical span)
    if target_key and target_key in all_months:
        ti = all_months.index(target_key)
        ax.axvspan(ti - 0.4, ti + 0.4, color=BLUE_LIGHT, alpha=0.5, zorder=1)

    # ── MoM % sticker on last valid data point ────────────────
    last_idx = next(
        (j for j in range(len(y_plot) - 1, -1, -1)
         if not np.isnan(y_plot[j])),
        None,
    )
    if last_idx is not None:
        last_val = y_plot[last_idx]

        if pct_val is None and last_idx > 0:
            prev = y_plot[last_idx - 1]
            if not np.isnan(prev) and prev > 0:
                pct_val = round((last_val - prev) / prev * 100, 1)

        # Circle marker
        ax.plot(last_idx, last_val, "o",
                color=ACCENT_PCT, markersize=13,
                markeredgecolor="white", markeredgewidth=1.0,
                zorder=5)

        if pct_val is not None:
            sign = "+" if pct_val > 0 else ""
            ax.annotate(
                f"{sign}{pct_val:.0f}%",
                xy=(last_idx, last_val),
                xytext=(0, 0),
                textcoords="offset points",
                fontsize=6, fontweight="bold",
                color="white",
                ha="center", va="center",
                zorder=6,
            )

    # X-axis: only first, middle, last label to avoid crowding
    n = len(all_months)
    show_idx = sorted({0, n // 2, n - 1})
    ax.set_xticks([x[i] for i in show_idx if i < n])
    ax.set_xticklabels([x_labels[i] for i in show_idx if i < n],
                       fontsize=5.5, color=TEXT_DARK)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax.tick_params(axis="y", colors=TEXT_DARK, labelsize=5.5)

    # Tenant name as title (truncate if > 20 chars)
    display = tenant_name if len(tenant_name) <= 20 else tenant_name[:18] + "…"
    ax.set_title(display, fontsize=8, fontweight="bold",
                 color=BLUE_DARK, pad=6)

    fig.tight_layout(pad=0.4)
    return _save(fig)
