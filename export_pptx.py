"""
Executive PowerPoint generator for Vir Ventures Weekly Analyst Report.
Theme: Midnight Navy + Burnt Orange — premium board-level presentation.
"""
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pandas as pd

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1A, 0x1F, 0x36)
ORANGE  = RGBColor(0xE8, 0x61, 0x1A)
ORANGE2 = RGBColor(0xF4, 0x87, 0x4B)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT   = RGBColor(0xF8, 0xF9, 0xFB)
GRAY    = RGBColor(0x6B, 0x72, 0x80)
LGRAY   = RGBColor(0xE5, 0xE7, 0xEB)
GREEN   = RGBColor(0x16, 0xA3, 0x4A)
RED     = RGBColor(0xDC, 0x26, 0x26)
DARK    = RGBColor(0x11, 0x18, 0x27)

W = Inches(13.33)   # LAYOUT_WIDE width
H = Inches(7.5)     # LAYOUT_WIDE height

def rgb(r, g, b): return RGBColor(r, g, b)

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_width=0):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, size=14, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txb

def add_kpi_card(slide, x, y, w, h, label, value, sub, accent_color=ORANGE):
    card = add_rect(slide, x, y, w, h, WHITE)
    card.shadow.inherit = False

    add_rect(slide, x, y, w, Inches(0.06), accent_color)

    add_text(slide, label, x + Inches(0.2), y + Inches(0.15),
             w - Inches(0.3), Inches(0.25), size=9, bold=True,
             color=GRAY, align=PP_ALIGN.LEFT)

    add_text(slide, value, x + Inches(0.2), y + Inches(0.42),
             w - Inches(0.3), Inches(0.7), size=32, bold=True,
             color=DARK, align=PP_ALIGN.LEFT)

    add_text(slide, sub, x + Inches(0.2), y + Inches(1.1),
             w - Inches(0.3), Inches(0.25), size=10,
             color=GRAY, align=PP_ALIGN.LEFT)

def make_chart_placeholder(slide, x, y, w, h, title, note="Chart generated in PowerPoint"):
    add_rect(slide, x, y, w, h, LIGHT)
    add_text(slide, title, x + Inches(0.2), y + Inches(0.15),
             w - Inches(0.4), Inches(0.3), size=11, bold=True, color=DARK)
    add_text(slide, note, x + Inches(0.2), y + h / 2 - Inches(0.2),
             w - Inches(0.4), Inches(0.35), size=10, color=GRAY, align=PP_ALIGN.CENTER)

# ── Slide builders ─────────────────────────────────────────────────────────────

def slide_cover(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, NAVY)

    # Orange accent block top-right
    add_rect(slide, W - Inches(3.2), Emu(0), Inches(3.2), Inches(2.2), ORANGE)
    add_rect(slide, W - Inches(2.0), Emu(0), Inches(2.0), Inches(3.6), RGBColor(0x25, 0x2B, 0x4A))

    # Company name
    add_text(slide, "VIR VENTURES", Inches(0.7), Inches(1.2),
             Inches(7), Inches(0.6), size=14, bold=True,
             color=ORANGE, align=PP_ALIGN.LEFT)

    # Title
    add_text(slide, "Weekly Analyst\nPerformance Report",
             Inches(0.7), Inches(2.0), Inches(8), Inches(2.0),
             size=44, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Week label + date
    add_text(slide, data["week_label"],
             Inches(0.7), Inches(4.15), Inches(6), Inches(0.45),
             size=16, color=ORANGE2, bold=False)

    add_text(slide, f"Report Date: {data['report_date']}  ·  Confidential — Internal Use Only",
             Inches(0.7), Inches(4.7), Inches(8), Inches(0.3),
             size=10, color=GRAY)

    # Bottom bar
    add_rect(slide, Emu(0), H - Inches(0.55), W, Inches(0.55), ORANGE)
    add_text(slide, "ANALYST PERFORMANCE TRACKER  ·  EXECUTIVE PRESENTATION",
             Inches(0.4), H - Inches(0.48), W - Inches(0.8), Inches(0.4),
             size=9, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    slide.notes_slide.notes_text_frame.text = (
        "Cover slide. Present this as the opening slide. "
        "Briefly introduce the weekly cadence and purpose of the report."
    )


def slide_executive_summary(prs, data):
    kpis = data["kpis"]
    a_df = data["analyst_df"]
    brand_df = data["brand_df"]
    bc_col = data["bc_col"]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, LIGHT)

    # Header bar
    add_rect(slide, Emu(0), Emu(0), W, Inches(0.85), NAVY)
    add_text(slide, "EXECUTIVE SUMMARY", Inches(0.5), Inches(0.15),
             Inches(8), Inches(0.55), size=11, bold=True, color=ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, data["week_label"], W - Inches(3.5), Inches(0.15),
             Inches(3.2), Inches(0.55), size=11, color=LGRAY, align=PP_ALIGN.RIGHT)

    # KPI row
    kpi_items = [
        ("CATALOGUE SIZE", f"{kpis['total']:,}", "Total SKUs reviewed", ORANGE),
        ("RECOMMENDED",    f"{kpis['recommended']:,}", "Yes + Yes, low qty", GREEN),
        ("REJECTED",       f"{kpis['rejected']:,}", "Flagged No", RED),
        ("APPROVAL RATE",  f"{kpis['rate']}%", f"of {kpis['total']:,} items", RGBColor(0x25, 0x63, 0xEB)),
    ]
    card_w = Inches(2.9)
    card_h = Inches(1.55)
    gap = Inches(0.22)
    start_x = Inches(0.4)
    card_y = Inches(1.05)

    for i, (label, val, sub, color) in enumerate(kpi_items):
        cx = start_x + i * (card_w + gap)
        add_kpi_card(slide, cx, card_y, card_w, card_h, label, val, sub, color)

    # Top analyst highlight
    if not a_df.empty:
        top = a_df.iloc[0]
        add_rect(slide, Inches(0.4), Inches(2.85), Inches(5.8), Inches(1.25), WHITE)
        add_text(slide, "★  TOP ANALYST THIS WEEK",
                 Inches(0.6), Inches(2.95), Inches(4), Inches(0.3),
                 size=9, bold=True, color=ORANGE)
        add_text(slide, str(top['Analyst']),
                 Inches(0.6), Inches(3.28), Inches(3.5), Inches(0.55),
                 size=24, bold=True, color=DARK)
        add_text(slide,
                 f"{int(top['Total Audited'])} audited  ·  {top['Approval Rate %']:.1f}% approval rate",
                 Inches(0.6), Inches(3.82), Inches(5), Inches(0.3),
                 size=11, color=GRAY)

    # Top brand highlight
    if not brand_df.empty and bc_col:
        top_b = brand_df.iloc[0]
        add_rect(slide, Inches(6.5), Inches(2.85), Inches(6.4), Inches(1.25), WHITE)
        add_text(slide, "★  TOP BRAND BY APPROVALS",
                 Inches(6.7), Inches(2.95), Inches(5), Inches(0.3),
                 size=9, bold=True, color=ORANGE)
        add_text(slide, str(top_b[bc_col]) if bc_col in top_b else str(top_b.iloc[0]),
                 Inches(6.7), Inches(3.28), Inches(5.5), Inches(0.55),
                 size=24, bold=True, color=DARK)
        add_text(slide, f"{int(top_b['Approved SKUs'])} approved SKUs",
                 Inches(6.7), Inches(3.82), Inches(5), Inches(0.3),
                 size=11, color=GRAY)

    # Bullet summary
    add_rect(slide, Inches(0.4), Inches(4.35), Inches(12.5), Inches(2.75), WHITE)
    add_text(slide, "KEY INSIGHTS", Inches(0.6), Inches(4.48),
             Inches(5), Inches(0.3), size=9, bold=True, color=ORANGE)

    top_analyst_name = a_df.iloc[0]['Analyst'] if not a_df.empty else "—"
    top_analyst_rate = a_df.iloc[0]['Approval Rate %'] if not a_df.empty else 0
    top_analyst_count = int(a_df.iloc[0]['Total Audited']) if not a_df.empty else 0
    top_brand_name = (brand_df.iloc[0][bc_col] if (not brand_df.empty and bc_col and bc_col in brand_df.columns)
                      else (str(brand_df.iloc[0].iloc[0]) if not brand_df.empty else "—"))
    top_brand_skus = int(brand_df.iloc[0]["Approved SKUs"]) if not brand_df.empty else 0

    bullets = [
        f"A total of {kpis['total']:,} SKUs were reviewed this week with an overall approval rate of {kpis['rate']}%.",
        f"{top_analyst_name} led analyst performance with {top_analyst_count} items audited ({top_analyst_rate:.1f}% approval rate).",
        f"{top_brand_name} is the highest-approved brand with {top_brand_skus} recommended SKUs — priority for sourcing.",
        f"{kpis['rejected']:,} SKUs were rejected and require re-evaluation before catalogue inclusion.",
    ]
    bullet_y = Inches(4.85)
    for b in bullets:
        add_text(slide, f"▸  {b}", Inches(0.6), bullet_y,
                 Inches(12.1), Inches(0.38), size=11, color=DARK)
        bullet_y += Inches(0.45)

    slide.notes_slide.notes_text_frame.text = (
        "Executive summary page. Walk through KPIs first, then highlight the top analyst "
        "and top brand. Use key insights as talking points."
    )


def slide_analyst_performance(prs, data):
    a_df = data["analyst_df"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, LIGHT)

    add_rect(slide, Emu(0), Emu(0), W, Inches(0.85), NAVY)
    add_text(slide, "ANALYST PERFORMANCE", Inches(0.5), Inches(0.15),
             Inches(8), Inches(0.55), size=11, bold=True, color=ORANGE)
    add_text(slide, data["week_label"], W - Inches(3.5), Inches(0.15),
             Inches(3.2), Inches(0.55), size=11, color=LGRAY, align=PP_ALIGN.RIGHT)
    add_text(slide, "Individual analyst audit quality and throughput breakdown",
             Inches(0.5), Inches(0.9), Inches(10), Inches(0.35), size=12, color=GRAY)

    if a_df.empty:
        add_text(slide, "No analyst data available.", Inches(2), Inches(3),
                 Inches(9), Inches(1), size=16, color=GRAY, align=PP_ALIGN.CENTER)
        return

    # Table
    table_x, table_y = Inches(0.4), Inches(1.45)
    col_widths = [Inches(2.5), Inches(1.8), Inches(1.8), Inches(1.8), Inches(2.2)]
    headers = ["Analyst", "Total Audited", "Recommended", "Rejected", "Approval Rate"]

    # Header row
    row_h = Inches(0.42)
    rx = table_x
    for i, (hdr, cw) in enumerate(zip(headers, col_widths)):
        add_rect(slide, rx, table_y, cw, row_h, NAVY)
        add_text(slide, hdr, rx + Inches(0.12), table_y + Inches(0.1),
                 cw - Inches(0.15), row_h - Inches(0.08), size=10, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
        rx += cw

    # Data rows
    for ri, (_, row) in enumerate(a_df.iterrows()):
        ry = table_y + row_h * (ri + 1)
        rx = table_x
        row_bg = WHITE if ri % 2 == 0 else LIGHT
        rate = row["Approval Rate %"]
        rate_color = GREEN if rate >= 70 else (RGBColor(0xF5, 0x9E, 0x0B) if rate >= 40 else RED)

        vals = [
            str(row["Analyst"]),
            str(int(row["Total Audited"])),
            str(int(row["Recommended"])),
            str(int(row["Rejected"])),
            f"{rate:.1f}%",
        ]
        for i, (val, cw) in enumerate(zip(vals, col_widths)):
            add_rect(slide, rx, ry, cw, row_h, row_bg)
            tc = rate_color if i == 4 else DARK
            bold = (i == 4)
            add_text(slide, val, rx + Inches(0.12), ry + Inches(0.1),
                     cw - Inches(0.15), row_h - Inches(0.08), size=11,
                     bold=bold, color=tc, align=PP_ALIGN.CENTER)
            rx += cw

    # Score cards on right side
    if len(a_df) > 0:
        sc_x = Inches(10.5)
        sc_items = [
            ("AVG APPROVAL", f"{a_df['Approval Rate %'].mean():.1f}%", "team average"),
            ("TOTAL AUDITED", f"{int(a_df['Total Audited'].sum()):,}", "combined"),
            ("TOP ANALYST", str(a_df.iloc[0]['Analyst']), f"{a_df.iloc[0]['Approval Rate %']:.0f}% rate"),
        ]
        sy = Inches(1.45)
        for label, val, sub in sc_items:
            add_rect(slide, sc_x, sy, Inches(2.45), Inches(1.4), WHITE)
            add_rect(slide, sc_x, sy, Inches(2.45), Inches(0.05), ORANGE)
            add_text(slide, label, sc_x + Inches(0.15), sy + Inches(0.15),
                     Inches(2.1), Inches(0.25), size=8, bold=True, color=GRAY)
            add_text(slide, val, sc_x + Inches(0.15), sy + Inches(0.42),
                     Inches(2.1), Inches(0.6), size=22, bold=True, color=DARK)
            add_text(slide, sub, sc_x + Inches(0.15), sy + Inches(1.05),
                     Inches(2.1), Inches(0.25), size=9, color=GRAY)
            sy += Inches(1.6)

    slide.notes_slide.notes_text_frame.text = (
        "Analyst performance page. Green = 70%+ approval, Amber = 40-69%, Red = below 40%. "
        "Discuss throughput vs quality balance with the team."
    )


def slide_brand_analysis(prs, data):
    brand_df = data["brand_df"]
    bc_col = data["bc_col"]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, LIGHT)

    add_rect(slide, Emu(0), Emu(0), W, Inches(0.85), NAVY)
    add_text(slide, "BRAND PERFORMANCE ANALYSIS", Inches(0.5), Inches(0.15),
             Inches(8), Inches(0.55), size=11, bold=True, color=ORANGE)
    add_text(slide, data["week_label"], W - Inches(3.5), Inches(0.15),
             Inches(3.2), Inches(0.55), size=11, color=LGRAY, align=PP_ALIGN.RIGHT)
    add_text(slide, "Top brands ranked by approved SKU count this week",
             Inches(0.5), Inches(0.9), Inches(10), Inches(0.35), size=12, color=GRAY)

    if brand_df.empty or not bc_col:
        add_text(slide, "No brand data available.", Inches(2), Inches(3),
                 Inches(9), Inches(1), size=16, color=GRAY, align=PP_ALIGN.CENTER)
        return

    name_col = bc_col if bc_col in brand_df.columns else brand_df.columns[0]
    max_val = brand_df["Approved SKUs"].max()
    bar_area_w = Inches(8.5)
    bar_h = Inches(0.44)
    bar_gap = Inches(0.12)
    start_x = Inches(3.0)
    start_y = Inches(1.45)
    label_x = Inches(0.4)

    for i, (_, row) in enumerate(brand_df.head(12).iterrows()):
        by = start_y + i * (bar_h + bar_gap)
        brand_name = str(row[name_col])
        count = int(row["Approved SKUs"])
        fill_w = bar_area_w * (count / max_val)

        # Brand label
        add_text(slide, brand_name, label_x, by + Inches(0.1),
                 Inches(2.4), bar_h, size=10, color=DARK, align=PP_ALIGN.RIGHT)

        # Background track
        add_rect(slide, start_x, by, bar_area_w, bar_h, LGRAY)

        # Fill bar — gradient effect via two bars
        bar_color = ORANGE if i == 0 else ORANGE2 if i < 3 else RGBColor(0xF9, 0xC0, 0x8E)
        add_rect(slide, start_x, by, fill_w, bar_h, bar_color)

        # Count label
        add_text(slide, str(count), start_x + fill_w + Inches(0.1), by + Inches(0.1),
                 Inches(0.6), bar_h, size=10, bold=(i == 0), color=DARK)

    slide.notes_slide.notes_text_frame.text = (
        "Brand performance chart. The top brand should receive priority in sourcing discussions. "
        "Discuss brands ranked 2-5 as pipeline candidates."
    )


def slide_recommendations_breakdown(prs, data):
    kpis = data["kpis"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, LIGHT)

    add_rect(slide, Emu(0), Emu(0), W, Inches(0.85), NAVY)
    add_text(slide, "SKU RECOMMENDATION BREAKDOWN", Inches(0.5), Inches(0.15),
             Inches(8), Inches(0.55), size=11, bold=True, color=ORANGE)
    add_text(slide, data["week_label"], W - Inches(3.5), Inches(0.15),
             Inches(3.2), Inches(0.55), size=11, color=LGRAY, align=PP_ALIGN.RIGHT)

    total = kpis["total"]
    rec   = kpis["recommended"]
    rej   = kpis["rejected"]
    other = kpis["other"]

    # Big donut representation using shapes
    cx = Inches(4.2)
    items = [
        ("Recommended", rec, GREEN, f"{rec/total*100:.1f}%"),
        ("Rejected",    rej, RED,   f"{rej/total*100:.1f}%"),
        ("Other",       other, LGRAY, f"{other/total*100:.1f}%"),
    ]

    # Visual cards side by side
    card_w = Inches(3.6)
    card_h = Inches(2.4)
    positions = [Inches(0.5), Inches(4.7), Inches(8.9)]
    for i, (label, count, color, pct) in enumerate(items):
        cx_card = positions[i]
        add_rect(slide, cx_card, Inches(1.3), card_w, card_h, WHITE)
        add_rect(slide, cx_card, Inches(1.3), card_w, Inches(0.08), color)
        add_text(slide, pct, cx_card + Inches(0.25), Inches(1.55),
                 card_w - Inches(0.4), Inches(0.85), size=48, bold=True, color=color)
        add_text(slide, label.upper(), cx_card + Inches(0.25), Inches(2.5),
                 card_w - Inches(0.4), Inches(0.3), size=9, bold=True, color=GRAY)
        add_text(slide, f"{count:,} SKUs", cx_card + Inches(0.25), Inches(2.85),
                 card_w - Inches(0.4), Inches(0.35), size=16, bold=True, color=DARK)

    # Horizontal stacked bar
    bar_y = Inches(4.1)
    bar_h2 = Inches(0.7)
    bar_x = Inches(0.5)
    bar_total_w = Inches(12.3)

    add_text(slide, "APPROVAL COMPOSITION", bar_x, bar_y - Inches(0.35),
             Inches(6), Inches(0.3), size=9, bold=True, color=GRAY)

    segments = [(rec, GREEN), (rej, RED), (other, LGRAY)]
    sx = bar_x
    for count, color in segments:
        seg_w = bar_total_w * (count / total) if total else 0
        if seg_w > 0:
            add_rect(slide, sx, bar_y, seg_w, bar_h2, color)
            if seg_w > Inches(0.5):
                add_text(slide, f"{count:,}", sx + Inches(0.05), bar_y + Inches(0.2),
                         seg_w - Inches(0.1), Inches(0.3), size=10, bold=True,
                         color=WHITE, align=PP_ALIGN.CENTER)
            sx += seg_w

    # Legend
    legend_items = [("Recommended", GREEN), ("Rejected", RED), ("Other / Blank", LGRAY)]
    lx = Inches(0.5)
    for label, color in legend_items:
        add_rect(slide, lx, bar_y + bar_h2 + Inches(0.15), Inches(0.2), Inches(0.2), color)
        add_text(slide, label, lx + Inches(0.28), bar_y + bar_h2 + Inches(0.12),
                 Inches(1.6), Inches(0.25), size=10, color=GRAY)
        lx += Inches(2.2)

    slide.notes_slide.notes_text_frame.text = (
        "SKU breakdown page. Focus on the approval rate and the composition bar. "
        "Rejected items should be reviewed in the next cycle."
    )


def slide_closing(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, NAVY)

    add_rect(slide, Emu(0), Emu(0), Inches(5), H, ORANGE)

    add_text(slide, "NEXT\nSTEPS", Inches(0.5), Inches(1.5),
             Inches(4), Inches(2.5), size=52, bold=True, color=WHITE)

    add_text(slide, "virventures.com", Inches(0.5), H - Inches(0.9),
             Inches(4), Inches(0.4), size=11, color=WHITE)

    steps = [
        "01  Prioritise sourcing for top-approved brands",
        "02  Re-evaluate rejected SKUs in next audit cycle",
        "03  Review analyst workload balance across the team",
        "04  Track week-on-week approval rate trend",
        "05  Share brand performance with vendor managers",
    ]
    sy = Inches(1.5)
    for step in steps:
        add_rect(slide, Inches(5.5), sy, Inches(7.3), Inches(0.75), RGBColor(0x25, 0x2B, 0x4A))
        add_text(slide, step, Inches(5.7), sy + Inches(0.18),
                 Inches(7), Inches(0.4), size=13, color=WHITE)
        sy += Inches(0.95)

    add_text(slide, f"Report prepared for management review  ·  {data['week_label']}",
             Inches(5.5), H - Inches(0.9), Inches(7.3), Inches(0.35),
             size=10, color=GRAY)

    slide.notes_slide.notes_text_frame.text = (
        "Closing slide. Walk through next steps with the management team. "
        "Confirm ownership for each action item before the meeting ends."
    )


# ── Main entry ─────────────────────────────────────────────────────────────────
def generate_pptx(data: dict) -> bytes:
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    slide_cover(prs, data)
    slide_executive_summary(prs, data)
    slide_analyst_performance(prs, data)
    slide_brand_analysis(prs, data)
    slide_recommendations_breakdown(prs, data)
    slide_closing(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
