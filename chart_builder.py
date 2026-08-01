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
TEXT_DARK   = "#2C3E50"
GRID_COLOR  = "#E5EAF0"
BG_WHITE    = "#FFFFFF"

MONTH_SHORT = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


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


# ── Chart 1: Monthly Sales Bar Chart ──────────────────────────

def chart_monthly_sales(monthly_totals: dict, target_key: str) -> io.BytesIO:
    """
    Bar chart: total sales per month across all tenants.

    monthly_totals = {"2026-01": 5_000_000, "2026-02": 6_200_000, ...}
    target_key     = "2026-05"
    """
    months = sorted(monthly_totals.keys())
    values = [monthly_totals[m] for m in months]
    labels = [
        f"{MONTH_SHORT[int(m.split('-')[1])]}-{m.split('-')[0][2:]}"
        for m in months
    ]

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

    # Highlight legend
    from matplotlib.patches import Patch
    legend = [
        Patch(color=BLUE_MID, label="Other months"),
        Patch(color=ACCENT_TEAL, label="Target month"),
    ]
    ax.legend(handles=legend, fontsize=8, framealpha=0.7,
              loc="upper left")

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

    tenant_sales = {"Tenant A": 5_000_000, "Tenant B": 3_200_000, ...}
    """
    # Sort descending, take top N
    sorted_items = sorted(tenant_sales.items(), key=lambda x: x[1], reverse=True)
    items = sorted_items[:top_n]

    if not items:
        return None

    names  = [x[0] for x in reversed(items)]   # bottom → top
    values = [x[1] for x in reversed(items)]

    # Color gradient: top bar gets accent, rest get blue shades
    n = len(items)
    colors = []
    for i in range(n):
        # i=n-1 is the top (highest value)
        if i == n - 1:
            colors.append(ACCENT_TEAL)
        elif i == n - 2:
            colors.append(BLUE_MID)
        else:
            # Fade toward lighter blue
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

    # Value labels at end of each bar
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

    target_label = (
        f"{MONTH_SHORT[int(target_key.split('-')[1])]}-"
        f"{target_key.split('-')[0][2:]}"
    )
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
    Dual-axis chart:
      Left axis  — visitor count (bar)
      Right axis — sales per visitor (line)
    """
    months = sorted(set(list(traffic_monthly.keys()) + list(sales_monthly.keys())))
    if not months:
        return None

    labels = [
        f"{MONTH_SHORT[int(m.split('-')[1])]}-{m.split('-')[0][2:]}"
        for m in months
    ]
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

    # Bar: traffic
    bar_colors = [BLUE_LIGHT if m != target_key else BLUE_MID for m in months]
    bars = ax1.bar(x, traffic_vals, color=bar_colors, width=0.55,
                   zorder=3, edgecolor="white", linewidth=0.5, label="Visitors")

    ax1.set_ylabel("Visitor Count", fontsize=9, color=BLUE_DARK)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax1.tick_params(axis="y", colors=BLUE_DARK)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)

    # Line: sales per visitor
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

    # Combined legend
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
    Includes traffic as a secondary axis if provided.

    daily_data    = [{"date": "2026-05-01", "sales": 12345}, ...]
    traffic_daily = [{"date": "2026-05-01", "traffic": 28066}, ...]
    """
    # Filter to target month
    filtered = [d for d in daily_data if d["date"].startswith(target_key)]
    if not filtered:
        return None

    # Aggregate sales by date (multiple tenants may share same date)
    from collections import defaultdict
    by_date = defaultdict(float)
    for d in filtered:
        by_date[d["date"]] += d["sales"]

    dates  = sorted(by_date.keys())
    values = [by_date[d] for d in dates]
    x_labels = [d[8:10] for d in dates]  # just day number "01", "02"...

    # Build traffic lookup
    traffic_vals = None
    if traffic_daily:
        traffic_by_date = defaultdict(float)
        for d in traffic_daily:
            if d["date"].startswith(target_key):
                traffic_by_date[d["date"]] += d["traffic"]
        if traffic_by_date:
            traffic_vals = [traffic_by_date.get(d, 0) for d in dates]

    # Detect weekends
    from datetime import datetime as dt_cls
    weekend_mask = [dt_cls.strptime(d, "%Y-%m-%d").weekday() >= 5 for d in dates]

    fig, ax1 = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor(BG_WHITE)
    ax1.set_facecolor(BG_WHITE)

    x = np.arange(len(dates))

    # Shade weekend columns
    for i, is_we in enumerate(weekend_mask):
        if is_we:
            ax1.axvspan(i - 0.5, i + 0.5, color="#E8F5E9", zorder=0, alpha=0.7)

    # Sales line (left axis)
    ax1.plot(x, values, color=BLUE_MID, linewidth=2.2, zorder=4, label="Sales")
    ax1.fill_between(x, values, alpha=0.12, color=BLUE_MID, zorder=3)

    # Dots: weekends in green, weekdays in blue
    for i, (val, is_we) in enumerate(zip(values, weekend_mask)):
        color = ACCENT_TEAL if is_we else BLUE_MID
        ax1.plot(i, val, "o", color=color, markersize=5, zorder=5)

    ax1.set_ylabel("Sales (IDR)", fontsize=9, color=BLUE_MID)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
    ax1.tick_params(axis="y", colors=BLUE_MID)
    ax1.spines["top"].set_visible(False)
    ax1.grid(axis="y", color=GRID_COLOR, linewidth=0.8, zorder=0)

    # Traffic line (right axis)
    if traffic_vals:
        ax2 = ax1.twinx()
        ax2.plot(x, traffic_vals, color=ACCENT_AMB, linewidth=2,
                 marker="s", markersize=4, zorder=5, alpha=0.85,
                 label="Traffic")
        ax2.set_ylabel("Visitor Count", fontsize=9, color=ACCENT_AMB)
        ax2.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_idr))
        ax2.tick_params(axis="y", colors=ACCENT_AMB)
        ax2.spines["top"].set_visible(False)

        # Combined legend
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

    # Mark highest sales day
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
