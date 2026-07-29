# pptx_builder.py
"""
Assembles the monthly management PowerPoint.
Uses python-pptx with a programmatic theme (no external template file needed).
"""
import io
import calendar as cal
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from collections import defaultdict


# ── Slide dimensions (16:9 widescreen) ────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── Brand colours ──────────────────────────────────────────────
C_NAVY      = RGBColor(0x1A, 0x3A, 0x5C)
C_BLUE      = RGBColor(0x28, 0x74, 0xA6)
C_TEAL      = RGBColor(0x1A, 0xBC, 0x9C)
C_AMBER     = RGBColor(0xF3, 0x9C, 0x12)
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT     = RGBColor(0xEC, 0xF0, 0xF1)
C_TEXT      = RGBColor(0x2C, 0x3E, 0x50)
C_MUTED     = RGBColor(0x7F, 0x8C, 0x8D)
C_RED       = RGBColor(0xE7, 0x4C, 0x3C)
C_GREEN     = RGBColor(0x27, 0xAE, 0x60)


def _rgb_to_hex(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


# ── Low-level helpers ──────────────────────────────────────────

def _add_rect(slide, left, top, width, height, fill_color, alpha=None):
    """Add a solid-filled rectangle."""
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height,
    )
    shape.line.fill.background()        # no border
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    return shape


def _add_textbox(
    slide, text, left, top, width, height,
    font_size=12, bold=False, color=C_TEXT,
    align=PP_ALIGN.LEFT, wrap=True, italic=False,
):
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size    = Pt(font_size)
    run.font.bold    = bold
    run.font.italic  = italic
    run.font.color.rgb = color
    run.font.name    = "Calibri"
    return txb


def _add_image(slide, img_buf, left, top, width, height):
    """Add a chart image (BytesIO) to the slide."""
    if img_buf is None:
        return
    img_buf.seek(0)
    slide.shapes.add_picture(img_buf, left, top, width, height)


def _add_bullet_textbox(
    slide, items: list[str], left, top, width, height,
    font_size=11, color=C_TEXT, bullet_char="▸",
):
    """Multi-line bullet textbox."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf  = txb.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"{bullet_char}  {item}"
        run.font.size  = Pt(font_size)
        run.font.color.rgb = color
        run.font.name  = "Calibri"
        # Paragraph spacing
        from pptx.util import Pt as PtU
        p.space_before = PtU(4)


def _header_bar(slide, title: str, subtitle: str = ""):
    """Dark navy header bar at top of every content slide."""
    _add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), C_NAVY)

    # Title
    _add_textbox(
        slide, title,
        left=Inches(0.4), top=Inches(0.1),
        width=Inches(10), height=Inches(0.65),
        font_size=22, bold=True, color=C_WHITE,
    )

    # Subtitle / month label
    if subtitle:
        _add_textbox(
            slide, subtitle,
            left=Inches(0.4), top=Inches(0.72),
            width=Inches(9), height=Inches(0.35),
            font_size=11, color=RGBColor(0xA9, 0xCC, 0xE3),
        )

    # Teal accent strip under header
    _add_rect(slide, 0, Inches(1.1), SLIDE_W, Inches(0.045), C_TEAL)


def _kpi_box(slide, label, value, left, top, width=Inches(2.5), height=Inches(1.2),
             value_color=C_NAVY, bg_color=C_LIGHT):
    """Single KPI card."""
    _add_rect(slide, left, top, width, height, bg_color)

    # Value (large)
    _add_textbox(
        slide, str(value),
        left=left + Inches(0.1), top=top + Inches(0.1),
        width=width - Inches(0.2), height=height * 0.55,
        font_size=18, bold=True, color=value_color,
        align=PP_ALIGN.CENTER,
    )
    # Label (small)
    _add_textbox(
        slide, label,
        left=left + Inches(0.1), top=top + height * 0.58,
        width=width - Inches(0.2), height=height * 0.42,
        font_size=9, color=C_MUTED,
        align=PP_ALIGN.CENTER,
    )


# ── Slide builders ─────────────────────────────────────────────

def _slide_cover(prs, month_label: str, generated_at: str):
    """Slide 1 — Cover."""
    layout = prs.slide_layouts[6]   # blank
    slide  = prs.slides.add_slide(layout)

    # Full-slide navy background
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_NAVY)

    # Teal accent bar (left edge)
    _add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, C_TEAL)

    # Mall name / logo placeholder
    _add_textbox(
        slide, "🏬  MALL MANAGEMENT",
        left=Inches(0.5), top=Inches(1.5),
        width=Inches(10), height=Inches(0.7),
        font_size=14, color=RGBColor(0xA9, 0xCC, 0xE3),
        bold=True,
    )

    # Main title
    _add_textbox(
        slide, "Monthly Performance Report",
        left=Inches(0.5), top=Inches(2.2),
        width=Inches(11), height=Inches(1.2),
        font_size=40, bold=True, color=C_WHITE,
    )

    # Month label
    _add_textbox(
        slide, month_label,
        left=Inches(0.5), top=Inches(3.4),
        width=Inches(8), height=Inches(0.8),
        font_size=28, color=C_TEAL, bold=True,
    )

    # Divider
    _add_rect(slide, Inches(0.5), Inches(4.25), Inches(5), Inches(0.04), C_TEAL)

    # Generated timestamp
    _add_textbox(
        slide, f"Generated: {generated_at}",
        left=Inches(0.5), top=Inches(4.4),
        width=Inches(8), height=Inches(0.4),
        font_size=10, color=C_MUTED,
    )

    # Confidential note
    _add_textbox(
        slide, "CONFIDENTIAL — Management Use Only",
        left=Inches(0.5), top=Inches(6.8),
        width=Inches(10), height=Inches(0.4),
        font_size=9, color=C_MUTED, italic=True,
    )


def _slide_executive_summary(prs, kpis: dict, llm_text: dict, month_label: str):
    """Slide 2 — Executive Summary."""
    layout = prs.slide_layouts[6]
    slide  = prs.slides.add_slide(layout)
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_WHITE)
    _header_bar(slide, "Executive Summary", month_label)

    def _fmt_idr(v):
        if v >= 1_000_000_000: return f"IDR {v/1_000_000_000:.2f}B"
        if v >= 1_000_000:     return f"IDR {v/1_000_000:.1f}M"
        return f"IDR {v:,.0f}"

    # ── KPI row ────────────────────────────────────────────────
    kpi_top = Inches(1.35)
    kpi_w   = Inches(2.4)
    kpi_h   = Inches(1.3)
    gap     = Inches(0.22)

    # Total Sales
    growth = kpis.get("sales_growth_pct")
    if growth is not None:
        growth_label = f"Total Sales  ({growth:+.1f}% MoM)"
        v_color = C_GREEN if growth >= 0 else C_RED
    else:
        growth_label = "Total Sales"
        v_color = C_NAVY

    _kpi_box(slide, growth_label, _fmt_idr(kpis["total_sales"]),
             left=Inches(0.4), top=kpi_top,
             width=kpi_w, height=kpi_h, value_color=v_color)

    # Visitors
    traffic = kpis["total_traffic"]
    _kpi_box(slide, "Total Visitors",
             f"{traffic:,.0f}" if traffic > 0 else "N/A",
             left=Inches(0.4) + kpi_w + gap, top=kpi_top,
             width=kpi_w, height=kpi_h)

    # Sales / Visitor
    spv = kpis.get("sales_per_visitor")
    _kpi_box(slide, "Sales / Visitor",
             _fmt_idr(spv) if spv else "N/A",
             left=Inches(0.4) + 2 * (kpi_w + gap), top=kpi_top,
             width=kpi_w, height=kpi_h)

    # Tenant Count
    _kpi_box(slide, "Active Tenants", str(kpis["total_tenants"]),
             left=Inches(0.4) + 3 * (kpi_w + gap), top=kpi_top,
             width=kpi_w, height=kpi_h)

    # Top Tenant
    _kpi_box(slide, f"Top: {kpis['top_tenant_name']}",
             _fmt_idr(kpis["top_tenant_sales"]),
             left=Inches(0.4) + 4 * (kpi_w + gap), top=kpi_top,
             width=kpi_w, height=kpi_h, value_color=C_TEAL)

    # ── Summary text ───────────────────────────────────────────
    _add_textbox(
        slide,
        llm_text.get("executive_summary", ""),
        left=Inches(0.4), top=Inches(2.85),
        width=Inches(12.4), height=Inches(2.6),
        font_size=12, color=C_TEXT, wrap=True,
    )

    # ── Bottom strip ───────────────────────────────────────────
    _add_rect(slide, 0, SLIDE_H - Inches(0.35), SLIDE_W, Inches(0.35), C_LIGHT)
    _add_textbox(
        slide, f"Data as of {month_label}  |  Tenant count: {kpis['total_tenants']}",
        left=Inches(0.3), top=SLIDE_H - Inches(0.33),
        width=Inches(10), height=Inches(0.3),
        font_size=8, color=C_MUTED,
    )


def _slide_with_chart(
    prs, title: str, subtitle: str,
    chart_buf: io.BytesIO,
    notes_text: str,
    chart_left=Inches(0.3), chart_top=Inches(1.25),
    chart_w=Inches(8.5), chart_h=Inches(5.0),
):
    """Generic slide: header + chart + notes panel."""
    layout = prs.slide_layouts[6]
    slide  = prs.slides.add_slide(layout)
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_WHITE)
    _header_bar(slide, title, subtitle)

    # Chart
    _add_image(slide, chart_buf,
               left=chart_left, top=chart_top,
               width=chart_w, height=chart_h)

    # Notes panel (right side)
    panel_left = chart_left + chart_w + Inches(0.2)
    panel_w    = SLIDE_W - panel_left - Inches(0.2)

    if panel_w > Inches(1.0):
        _add_rect(slide, panel_left, Inches(1.25),
                  panel_w, Inches(5.5),
                  RGBColor(0xF4, 0xF6, 0xF9))

        _add_textbox(
            slide, "Key Findings",
            left=panel_left + Inches(0.1), top=Inches(1.35),
            width=panel_w - Inches(0.2), height=Inches(0.4),
            font_size=10, bold=True, color=C_NAVY,
        )

        _add_textbox(
            slide, notes_text,
            left=panel_left + Inches(0.1), top=Inches(1.85),
            width=panel_w - Inches(0.2), height=Inches(4.8),
            font_size=10, color=C_TEXT, wrap=True,
        )

    # Slide number hint in bottom-right
    _add_rect(slide, 0, SLIDE_H - Inches(0.3), SLIDE_W, Inches(0.3), C_LIGHT)

    return slide

def _slide_events(prs, events_data: dict, target_key: str, month_label: str):
    """Slide: Event Calendar for the month."""
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_WHITE)
    _header_bar(slide, "Event Calendar", month_label)

    # Filter to target month
    target_events = [e for e in events_data.get("events_flat", [])
                     if e["date"].startswith(target_key)]

    if not target_events:
        _add_textbox(
            slide, "No events found for this period.",
            left=Inches(0.4), top=Inches(1.5),
            width=Inches(10), height=Inches(0.5),
            font_size=14, color=C_MUTED,
        )
        return

    # Group by date
    events_by_date = defaultdict(list)
    for e in target_events:
        try:
            dt = datetime.strptime(e["date"], "%Y-%m-%d")
            label = dt.strftime("%a, %d %b")
        except ValueError:
            label = e["date"]
        events_by_date[label].append(e)

    # Build a scrollable-looking list
    # Left column: date cards
    date_list = sorted(events_by_date.keys())

    # Title showing event count
    _add_textbox(
        slide,
        f"{len(target_events)} events across {len(date_list)} day(s)",
        left=Inches(0.4), top=Inches(1.3),
        width=Inches(10), height=Inches(0.35),
        font_size=11, color=C_MUTED, italic=True,
    )

    # Simplified layout: 2-column grid of event cards
    card_w = Inches(5.9)
    card_h = Inches(1.6)
    gap_x = Inches(0.35)
    gap_y = Inches(0.2)

    cols = 2
    start_x = Inches(0.3)
    start_y = Inches(1.75)

    for i, (date_label, evts) in enumerate(events_by_date.items()):
        col = i % cols
        row_idx = i // cols

        left = start_x + col * (card_w + gap_x)
        top = start_y + row_idx * (card_h + gap_y)

        if top + card_h > Inches(7.2):
            break  # Don't overflow the slide

        # Card background
        _add_rect(slide, left, top, card_w, card_h, RGBColor(0xF4, 0xF6, 0xF9))

        # Date header
        _add_textbox(
            slide, date_label,
            left=left + Inches(0.1), top=top + Inches(0.05),
            width=card_w - Inches(0.2), height=Inches(0.3),
            font_size=9, bold=True, color=C_NAVY,
        )

        # Event list
        event_text = "\n".join(
            f"• {e['event_name'][:80]}  ({e['location']})"
            for e in evts[:5]
        )
        if len(evts) > 5:
            event_text += f"\n... +{len(evts) - 5} more"

        _add_textbox(
            slide, event_text,
            left=left + Inches(0.1), top=top + Inches(0.38),
            width=card_w - Inches(0.2), height=card_h - Inches(0.5),
            font_size=8, color=C_TEXT, wrap=True,
        )

    # Bottom strip
    _add_rect(slide, 0, SLIDE_H - Inches(0.35), SLIDE_W, Inches(0.35), C_LIGHT)
    _add_textbox(
        slide, f"Event count: {len(target_events)}",
        left=Inches(0.3), top=SLIDE_H - Inches(0.33),
        width=Inches(10), height=Inches(0.3),
        font_size=8, color=C_MUTED,
    )
    
def _slide_recommendations(prs, recommendations: list[str], month_label: str):
    """Last slide — Recommendations."""
    layout = prs.slide_layouts[6]
    slide  = prs.slides.add_slide(layout)
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_NAVY)

    # Header
    _add_textbox(
        slide, "Recommendations & Next Steps",
        left=Inches(0.5), top=Inches(0.4),
        width=Inches(12), height=Inches(0.7),
        font_size=24, bold=True, color=C_WHITE,
    )
    _add_textbox(
        slide, month_label,
        left=Inches(0.5), top=Inches(1.1),
        width=Inches(8), height=Inches(0.4),
        font_size=12, color=C_TEAL,
    )
    _add_rect(slide, Inches(0.5), Inches(1.55), Inches(12.3), Inches(0.04), C_TEAL)

    # Recommendation cards
    card_colors = [
        RGBColor(0x1B, 0x4F, 0x72),
        RGBColor(0x1A, 0x53, 0x76),
        RGBColor(0x1A, 0x5C, 0x88),
        RGBColor(0x1F, 0x61, 0x8D),
    ]
    card_h = Inches(1.15)
    card_gap = Inches(0.18)
    card_top_start = Inches(1.75)

    icons = ["01", "02", "03", "04"]

    for i, rec in enumerate(recommendations[:4]):
        top   = card_top_start + i * (card_h + card_gap)
        color = card_colors[i % len(card_colors)]

        # Card background
        _add_rect(slide, Inches(0.5), top, Inches(12.3), card_h, color)

        # Number badge
        _add_textbox(
            slide, icons[i],
            left=Inches(0.6), top=top + Inches(0.25),
            width=Inches(0.6), height=Inches(0.65),
            font_size=22, bold=True, color=C_TEAL,
        )

        # Recommendation text
        _add_textbox(
            slide, rec,
            left=Inches(1.3), top=top + Inches(0.12),
            width=Inches(11.3), height=card_h - Inches(0.24),
            font_size=13, color=C_WHITE, wrap=True,
        )

    # Footer
    _add_textbox(
        slide, "This report was generated automatically. Data subject to final audit.",
        left=Inches(0.5), top=Inches(7.1),
        width=Inches(12), height=Inches(0.3),
        font_size=8, color=C_MUTED, italic=True,
    )


# ── Main assembler ─────────────────────────────────────────────

def build_pptx(
    kpis       : dict,
    llm_text   : dict,
    charts     : dict,        # {"monthly_sales": BytesIO, "top_tenants": BytesIO, ...}
    month_label: str,
    events_data=None,
) -> io.BytesIO:
    """
    Assemble the full presentation.

    charts keys:
        monthly_sales   — bar chart
        top_tenants     — horizontal bar
        traffic         — dual-axis line/bar
        daily_sales     — daily line chart
    """
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    generated_at = datetime.now().strftime("%d %B %Y, %H:%M")

    # ── Slide 1: Cover ────────────────────────────────────────
    _slide_cover(prs, month_label, generated_at)

    # ── Slide 2: Executive Summary ────────────────────────────
    _slide_executive_summary(prs, kpis, llm_text, month_label)

    # ── Slide 3: Total Monthly Sales ─────────────────────────
    if charts.get("monthly_sales"):
        _slide_with_chart(
            prs,
            title    = "Monthly Sales Performance",
            subtitle = month_label,
            chart_buf   = charts["monthly_sales"],
            notes_text  = llm_text.get("sales_slide_notes", ""),
            chart_left  = Inches(0.3),
            chart_top   = Inches(1.25),
            chart_w     = Inches(8.8),
            chart_h     = Inches(5.5),
        )

    # ── Slide 4: Top Tenants ──────────────────────────────────
    if charts.get("top_tenants"):
        _slide_with_chart(
            prs,
            title    = "Top Tenant Performance",
            subtitle = month_label,
            chart_buf   = charts["top_tenants"],
            notes_text  = llm_text.get("tenant_slide_notes", ""),
            chart_left  = Inches(0.3),
            chart_top   = Inches(1.25),
            chart_w     = Inches(8.8),
            chart_h     = Inches(5.5),
        )

    # ── Slide 5: Traffic ──────────────────────────────────────
    if charts.get("traffic"):
        _slide_with_chart(
            prs,
            title    = "Visitor Traffic & Spend Efficiency",
            subtitle = month_label,
            chart_buf   = charts["traffic"],
            notes_text  = llm_text.get("traffic_slide_notes", ""),
            chart_left  = Inches(0.3),
            chart_top   = Inches(1.25),
            chart_w     = Inches(8.8),
            chart_h     = Inches(5.5),
        )

    # ── Slide 6: Daily Sales ──────────────────────────────────
    if charts.get("daily_sales"):
        _slide_with_chart(
            prs,
            title    = "Daily Sales Pattern",
            subtitle = month_label,
            chart_buf   = charts["daily_sales"],
            notes_text  = llm_text.get("daily_slide_notes", ""),
            chart_left  = Inches(0.3),
            chart_top   = Inches(1.25),
            chart_w     = Inches(8.8),
            chart_h     = Inches(5.5),
        )

    # ── Slide 7: Events Calendar ──────────────────────────────
    if events_data and events_data.get("events_flat"):
        _slide_events(prs, events_data, None, month_label)

    # ── Slide 8: Recommendations ─────────────────────────────
    _slide_recommendations(
        prs,
        recommendations = llm_text.get("recommendations", []),
        month_label     = month_label,
    )

    # ── Save to BytesIO ───────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
