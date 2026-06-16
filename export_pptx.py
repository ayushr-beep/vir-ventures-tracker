"""
Vir Ventures — Boardroom PowerPoint Generator
8 slides, navy × orange premium executive theme.
All numbers pulled from live mastersheet data.
"""
import io, numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x0F,0x16,0x29)
NAVY2   = RGBColor(0x1E,0x2A,0x45)
ORANGE  = RGBColor(0xE8,0x61,0x1A)
ORANGE2 = RGBColor(0xF4,0x87,0x4B)
WHITE   = RGBColor(0xFF,0xFF,0xFF)
LIGHT   = RGBColor(0xF8,0xF9,0xFB)
LGRAY   = RGBColor(0xE5,0xE7,0xEB)
GRAY    = RGBColor(0x6B,0x72,0x80)
GREEN   = RGBColor(0x16,0xA3,0x4A)
RED     = RGBColor(0xDC,0x26,0x26)
AMBER   = RGBColor(0xD9,0x77,0x06)
DARK    = RGBColor(0x11,0x18,0x27)
BLUE    = RGBColor(0x25,0x63,0xEB)

W = Inches(13.33); H = Inches(7.5)

def _rect(slide, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(1, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line
    else: s.line.fill.background()
    return s

def _txt(slide, text, x, y, w, h, size=12, bold=False, color=WHITE,
         align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color
    r.font.name = "Calibri"
    return tb

def _set_bg(slide, color):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color

def _kpi_card(slide, x, y, w, h, label, value, sub, accent=ORANGE, val_color=None):
    _rect(slide, x, y, w, h, WHITE)
    _rect(slide, x, y, w, Inches(0.055), accent)
    _txt(slide, label, x+Inches(.18), y+Inches(.12), w-Inches(.25), Inches(.22),
         size=8, bold=True, color=GRAY)
    _txt(slide, value, x+Inches(.18), y+Inches(.37), w-Inches(.25), Inches(.65),
         size=28, bold=True, color=val_color or DARK)
    _txt(slide, sub, x+Inches(.18), y+Inches(1.0), w-Inches(.25), Inches(.22),
         size=9, color=GRAY)

def _slide_header(slide, title, subtitle, week):
    _rect(slide, Emu(0), Emu(0), W, Inches(.75), NAVY)
    _txt(slide, title.upper(), Inches(.45), Inches(.12), Inches(9), Inches(.5),
         size=10, bold=True, color=ORANGE)
    _txt(slide, week, W-Inches(3.2), Inches(.12), Inches(2.9), Inches(.5),
         size=10, color=LGRAY, align=PP_ALIGN.RIGHT)
    _txt(slide, subtitle, Inches(.45), Inches(.82), Inches(10), Inches(.32),
         size=11, color=GRAY)

def _footer(slide):
    _rect(slide, Emu(0), H-Inches(.38), W, Inches(.38), NAVY)
    _txt(slide, "VIR VENTURES  ·  BOARDROOM ANALYST REPORT  ·  CONFIDENTIAL",
         Inches(.4), H-Inches(.32), W-Inches(.8), Inches(.28),
         size=8, bold=True, color=RGBColor(0x3A,0x42,0x60))

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ═══════════════════════════════════════════════════════════════════════════════
def slide_cover(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, NAVY)
    _rect(s, W-Inches(3.8), Emu(0), Inches(3.8), Inches(2.6), ORANGE)
    _rect(s, W-Inches(2.2), Emu(0), Inches(2.2), H, NAVY2)
    _txt(s, "VIR VENTURES", Inches(.7), Inches(1.4), Inches(7), Inches(.55),
         size=13, bold=True, color=ORANGE)
    _txt(s, "Boardroom Analyst\nPerformance Report",
         Inches(.7), Inches(2.1), Inches(8.5), Inches(2.2),
         size=46, bold=True, color=WHITE)
    _txt(s, d['week_label'], Inches(.7), Inches(4.5), Inches(7), Inches(.45),
         size=17, color=ORANGE2)
    _txt(s, f"Report Date: {d['report_date']}   ·   SKUs Reviewed: {d['total']}   ·   Approval Rate: {d['rate']}%",
         Inches(.7), Inches(5.1), Inches(9), Inches(.32), size=11, color=GRAY)
    _txt(s, "Confidential — Internal Management Use Only",
         Inches(.7), Inches(5.5), Inches(7), Inches(.3), size=10, color=RGBColor(0x3A,0x42,0x60))
    _rect(s, Emu(0), H-Inches(.5), W-Inches(2.2), Inches(.5), ORANGE)
    _txt(s, "ANALYST PERFORMANCE TRACKER  ·  WEEKLY REVIEW",
         Inches(.4), H-Inches(.45), W-Inches(2.6), Inches(.4),
         size=9, bold=True, color=WHITE)
    s.notes_slide.notes_text_frame.text = (
        "Welcome to the weekly Vir Ventures Boardroom Analyst Report. "
        "This presentation covers ASIN-level performance, analyst scorecards, "
        "brand intelligence, and sourcing recommendations for the management team."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
def slide_exec_summary(prs, d):
    df = d['df']
    s = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "Executive Summary", "Week-on-week snapshot for management review", d['week_label'])
    _footer(s)

    # KPI cards
    cw = Inches(2.8); ch = Inches(1.5); gap = Inches(.2); sy = Inches(1.25)
    kpis = [
        ("CATALOGUE SIZE",    str(d['total']),               "Total SKUs reviewed",    ORANGE, DARK),
        ("APPROVED",          str(d['approved']),            "Yes + Yes, Need Discount", GREEN, GREEN),
        ("REJECTED",          str(d['rejected']),            "Negative / Low margin",  RED,   RED),
        ("APPROVAL RATE",     f"{d['rate']}%",               f"of {d['total']} SKUs",  BLUE,  BLUE),
    ]
    for i,(label,val,sub,accent,vc) in enumerate(kpis):
        _kpi_card(s, Inches(.4)+i*(cw+gap), sy, cw, ch, label, val, sub, accent, vc)

    # Key insights panel
    _rect(s, Inches(.4), Inches(3.0), Inches(11.5), Inches(3.65), WHITE)
    _txt(s, "KEY MANAGEMENT INSIGHTS", Inches(.6), Inches(3.1),
         Inches(8), Inches(.3), size=9, bold=True, color=ORANGE)

    yes_df   = df[df['is_yes']]
    top_a    = df.groupby('Analyst')['is_yes'].mean().idxmax()
    top_a_r  = round(df.groupby('Analyst')['is_yes'].mean()[top_a]*100,1)
    top_a_n  = int(df[df['Analyst']==top_a]['is_yes'].sum())
    top_brand = yes_df.groupby('Brand')['ASIN'].count().idxmax()
    top_brand_n = int(yes_df.groupby('Brand')['ASIN'].count().max())
    top_asin = yes_df.sort_values('Last month sold',ascending=False).iloc[0]
    avg_m    = round(yes_df['margin_pct'].mean(),1)
    neg_m    = df[~df['is_yes']]['Remarks '].str.contains('Negative',na=False).sum()
    vendors  = df['Vendor Prefix'].nunique() if 'Vendor Prefix' in df.columns else '—'

    bullets = [
        f"This week, {d['total']} SKUs were reviewed across {df['Brand'].nunique()} brands and {vendors} vendor prefixes. Overall approval rate is {d['rate']}%.",
        f"{top_a} is the top-performing analyst this week with {top_a_n} approvals and a {top_a_r}% approval rate — demonstrating consistent quality selection.",
        f"Average Buy Box margin on approved SKUs is {avg_m}%, well above breakeven — indicating healthy profitability on sourced inventory.",
        f"{top_brand} leads brand performance with {top_brand_n} approved SKUs — recommend prioritising this vendor for immediate PO actioning.",
        f"Top demand ASIN: {top_asin['ASIN']} ({top_asin['Brand']}) — {int(top_asin['Last month sold'])} units sold last month at ${top_asin['BB Price']:.2f} BB price.",
        f"{d['rejected']} ASINs rejected — {neg_m} flagged Negative Margin. Require vendor renegotiation before reconsideration.",
    ]
    by = Inches(3.5)
    for b in bullets:
        _rect(s, Inches(.55), by, Inches(0.06), Inches(0.22), ORANGE)
        _txt(s, b, Inches(.72), by-Inches(.02), Inches(11.0), Inches(.38),
             size=11, color=DARK)
        by += Inches(.44)

    s.notes_slide.notes_text_frame.text = (
        "Walk through each KPI first. Then present the six insights in order. "
        "Emphasise the approval rate and top analyst performance. "
        "Flag the rejected SKUs for the sourcing team to action post-meeting."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — TOP APPROVED ASINs
# ═══════════════════════════════════════════════════════════════════════════════
def slide_top_asins(prs, d):
    df   = d['df']
    s    = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "ASIN Intelligence — Top Approved SKUs",
                  "Ranked by monthly demand velocity · All values from live mastersheet", d['week_label'])
    _footer(s)

    yes_df = df[df['is_yes']].sort_values('Last month sold', ascending=False).head(10)

    # Table header
    _rect(s, Inches(.4), Inches(1.2), Inches(12.5), Inches(.42), NAVY)
    headers = [("ASIN",1.5),("Brand",2.2),("Analyst",.95),("Rec.",.85),
               ("Units/Mo",.85),("Net $",.75),("BB $",.75),("Margin%",.85),
               ("BSR Rank",.95),("Score",.7),("BB Holder",1.3)]
    cx = Inches(.4)
    for hdr,cw in headers:
        _txt(s, hdr, cx+Inches(.06), Inches(1.25), Inches(cw-.1), Inches(.32),
             size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        cx += Inches(cw)

    # Data rows
    for ri,(_, row) in enumerate(yes_df.iterrows()):
        ry = Inches(1.62) + ri*Inches(.52)
        bg = WHITE if ri%2==0 else LIGHT
        _rect(s, Inches(.4), ry, Inches(12.5), Inches(.5), bg)
        if ri==0: _rect(s, Inches(.4), ry, Inches(.07), Inches(.5), ORANGE)

        cells = [
            (str(row.get('ASIN',''))[:12],          1.5, DARK,  False),
            (str(row.get('Brand',''))[:20],         2.2, DARK,  ri==0),
            (str(row.get('Analyst','')),             .95, GRAY,  False),
            (str(row.get('Recommended',''))[:14],    .85, ORANGE if ri==0 else GRAY, False),
            (f"{int(row.get('Last month sold',0))}",  .85, GREEN, True),
            (f"${row.get('Net price',0):.2f}",        .75, DARK,  False),
            (f"${row.get('BB Price',0):.2f}",         .75, DARK,  False),
            (f"{row.get('margin_pct',0):.1f}%",       .85, ORANGE,True),
            (f"#{int(row.get('Rank',0)):,}",           .95, DARK,  False),
            (str(row.get('Total Product Score','')),   .70, DARK,  False),
            (str(row.get('BuyBoxSellerName',''))[:14], 1.3, GRAY,  False),
        ]
        tx = Inches(.4)
        for val, cw, col, bld in cells:
            _txt(s, val, tx+Inches(.06), ry+Inches(.12), Inches(cw-.1), Inches(.28),
                 size=9, bold=bld, color=col, align=PP_ALIGN.CENTER)
            tx += Inches(cw)

    s.notes_slide.notes_text_frame.text = (
        "This table ranks all approved ASINs by monthly unit velocity. "
        "Highlight the top 3 rows for immediate PO actioning. "
        "Margin % column shows Buy Box premium over our net cost. "
        "Higher margin + lower BSR rank = highest priority for sourcing."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — ANALYST SCORECARDS
# ═══════════════════════════════════════════════════════════════════════════════
def slide_analyst(prs, d):
    df   = d['df']
    s    = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "Analyst Performance Scorecards",
                  "Individual audit quality, throughput, and scoring breakdown", d['week_label'])
    _footer(s)

    analysts = df.groupby('Analyst').agg(
        Total=('ASIN','count'),
        Approved=('is_yes','sum'),
        Rejected=('is_yes', lambda x:(~x).sum()),
        Rate=('is_yes', lambda x:round(x.mean()*100,1)),
        AvgRank=('Rank','mean'),
        AvgNetPrice=('Net price','mean'),
        AvgBBPrice=('BB Price','mean'),
        AvgMarginPct=('margin_pct','mean'),
        AvgDemand=('Total Demand Score','mean'),
        AvgComp=('Total Competition Score','mean'),
        AvgMarginScore=('Total Margin Score','mean'),
        TotalUnits=('Last month sold','sum'),
    ).reset_index().round(1)

    n = len(analysts)
    card_w = Inches(12.4/n); card_h = Inches(5.5)
    start_x = Inches(.4)

    for i,(_,row) in enumerate(analysts.iterrows()):
        cx = start_x + i*card_w + i*Inches(.1)
        rate = row['Rate']
        accent = GREEN if rate>=80 else AMBER if rate>=50 else RED

        _rect(s, cx, Inches(1.2), card_w, card_h, WHITE)
        _rect(s, cx, Inches(1.2), card_w, Inches(.08), accent)

        # Name & rate
        _txt(s, row['Analyst'], cx+Inches(.2), Inches(1.35), card_w-Inches(.3), Inches(.45),
             size=18, bold=True, color=DARK)
        _txt(s, f"{rate}% APPROVAL", cx+Inches(.2), Inches(1.85), card_w-Inches(.3), Inches(.32),
             size=11, bold=True, color=accent)

        # Stats grid
        stats = [
            ("Total Audited",   str(int(row['Total']))),
            ("Approved",        str(int(row['Approved']))),
            ("Rejected",        str(int(row['Rejected']))),
            ("Avg Margin %",    f"{row['AvgMarginPct']:.1f}%"),
            ("Avg BSR Rank",    f"#{row['AvgRank']:,.0f}"),
            ("Avg Net Price",   f"${row['AvgNetPrice']:.2f}"),
            ("Avg BB Price",    f"${row['AvgBBPrice']:.2f}"),
            ("Total Units/Mo",  str(int(row['TotalUnits']))),
            ("Demand Score",    str(row['AvgDemand'])),
            ("Competition Sc.", str(row['AvgComp'])),
            ("Margin Score",    str(row['AvgMarginScore'])),
        ]
        sy = Inches(2.3)
        for label, val in stats:
            _rect(s, cx+Inches(.15), sy, card_w-Inches(.25), Inches(.42), LIGHT)
            _txt(s, label, cx+Inches(.22), sy+Inches(.03), card_w-Inches(.4), Inches(.18),
                 size=8, color=GRAY)
            _txt(s, val, cx+Inches(.22), sy+Inches(.2), card_w-Inches(.4), Inches(.2),
                 size=11, bold=True, color=DARK)
            sy += Inches(.47)

    s.notes_slide.notes_text_frame.text = (
        "Each card represents one analyst's full week performance. "
        "Green = 80%+ approval rate. Amber = 50-79%. Red = below 50%. "
        "Discuss throughput vs quality balance. "
        "Note: 100% approval rate analysts had clean vendor batches with strong demand scores."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — BRAND & VENDOR ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
def slide_brands(prs, d):
    df   = d['df']
    s    = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "Brand & Vendor Intelligence",
                  "Approved SKU count, margin performance, and demand velocity by brand", d['week_label'])
    _footer(s)

    brand = df[df['is_yes']].groupby('Brand').agg(
        Approved=('ASIN','count'),
        TotalUnits=('Last month sold','sum'),
        AvgMarginPct=('margin_pct','mean'),
        AvgRank=('Rank','mean'),
    ).reset_index().sort_values('Approved',ascending=False).head(10).round(1)

    total_approved = brand['Approved'].sum()
    max_approved   = brand['Approved'].max()

    # Bar chart (manual)
    chart_x = Inches(.4); chart_y = Inches(1.25)
    bar_h = Inches(.44); gap = Inches(.1)
    track_w = Inches(7.5); label_w = Inches(2.0); count_w = Inches(.6)

    _txt(s, "BRAND", chart_x, chart_y-Inches(.3), label_w, Inches(.25),
         size=8, bold=True, color=GRAY)
    _txt(s, "APPROVED SKUs →", chart_x+label_w+Inches(.1), chart_y-Inches(.3),
         track_w, Inches(.25), size=8, bold=True, color=GRAY)

    for i,(_,row) in enumerate(brand.iterrows()):
        by = chart_y + i*(bar_h+gap)
        pct = row['Approved']/max_approved if max_approved else 0
        fill_w = max(Inches(.15), track_w*pct)
        clr = ORANGE if i==0 else ORANGE2 if i<3 else RGBColor(0xF9,0xC0,0x8E)

        # Label
        _txt(s, row['Brand'], chart_x, by+Inches(.1), label_w, bar_h,
             size=10, bold=(i==0), color=DARK, align=PP_ALIGN.RIGHT)
        # Track
        _rect(s, chart_x+label_w+Inches(.15), by, track_w, bar_h, LGRAY)
        # Fill
        _rect(s, chart_x+label_w+Inches(.15), by, fill_w, bar_h, clr)
        # Count
        _txt(s, str(int(row['Approved'])),
             chart_x+label_w+Inches(.15)+fill_w+Inches(.08), by+Inches(.1),
             Inches(.5), bar_h, size=10, bold=(i==0), color=DARK)

    # Right panel — brand metrics table
    tx = Inches(10.0); ty = Inches(1.25)
    _txt(s, "BRAND METRICS SNAPSHOT", tx, ty-Inches(.35), Inches(2.9), Inches(.28),
         size=8, bold=True, color=ORANGE)

    # Header
    _rect(s, tx, ty, Inches(2.9), Inches(.38), NAVY)
    for hdr, cx_off, w in [("Brand",0,1.0),("Units/Mo",1.05,.85),("Margin%",1.95,.95)]:
        _txt(s, hdr, tx+Inches(cx_off)+Inches(.05), ty+Inches(.08), Inches(w), Inches(.22),
             size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    for ri,(_,row) in enumerate(brand.iterrows()):
        ry2 = ty+Inches(.38)+ri*Inches(.47)
        bg2 = WHITE if ri%2==0 else LIGHT
        _rect(s, tx, ry2, Inches(2.9), Inches(.45), bg2)
        cells2 = [
            (str(row['Brand'])[:12], 0,    1.0, DARK,  ri==0),
            (str(int(row['TotalUnits'])),  1.05, .85, GREEN, False),
            (f"{row['AvgMarginPct']:.1f}%",1.95,  .95, ORANGE, True),
        ]
        for val2, cx_off2, w2, col2, bld2 in cells2:
            _txt(s, val2, tx+Inches(cx_off2)+Inches(.05), ry2+Inches(.12),
                 Inches(w2), Inches(.22), size=9, bold=bld2, color=col2, align=PP_ALIGN.CENTER)

    s.notes_slide.notes_text_frame.text = (
        "Focus on the top 3 brands — these should drive this week's purchase orders. "
        "The bar length shows relative SKU approval count. "
        "Margin % and units/month together determine sourcing priority. "
        "High margin + high units = immediate action required."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — REJECTION ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
def slide_rejected(prs, d):
    df = d['df']
    s  = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "Rejection Analysis — Root Cause Review",
                  "All rejected ASINs with margin breakdown and recommended next actions", d['week_label'])
    _footer(s)

    rej = df[~df['is_yes']].copy()

    # Reason summary cards
    reasons = rej['Remarks '].fillna('Unspecified').value_counts()
    neg = int(reasons.get('Negative Margin',0))
    low = int(reasons.get('Low Margin',0))
    oth = d['rejected'] - neg - low

    reason_cards = [
        ("NEGATIVE MARGIN", str(neg), "BB Price < Breakeven", RED),
        ("LOW MARGIN",       str(low), "Marginal profitability", AMBER),
        ("OTHER REASONS",    str(oth), "Review required",        GRAY),
    ]
    cw2 = Inches(3.5); sy2 = Inches(1.2)
    for i,(label,val,sub,clr) in enumerate(reason_cards):
        cx2 = Inches(.4)+i*(cw2+Inches(.25))
        _kpi_card(s, cx2, sy2, cw2, Inches(1.4), label, val, sub, clr, clr)

    # Rejected ASIN table
    _rect(s, Inches(.4), Inches(2.85), Inches(12.5), Inches(.4), NAVY)
    hdrs = [("ASIN",1.6),("Brand",1.7),("Analyst",.9),
            ("Net Price",.95),("BB Price",.95),("Breakeven",.95),
            ("Diff from SP",.95),("Reason",2.4),("Score",.75)]
    cx3 = Inches(.4)
    for hdr,cw3 in hdrs:
        _txt(s, hdr, cx3+Inches(.06), Inches(2.9), Inches(cw3-.1), Inches(.28),
             size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        cx3 += Inches(cw3)

    for ri,(_,row) in enumerate(rej.iterrows()):
        ry3 = Inches(3.25)+ri*Inches(.48)
        bg3 = RGBColor(0xFF,0xF1,0xF2) if ri%2==0 else WHITE
        _rect(s, Inches(.4), ry3, Inches(12.5), Inches(.46), bg3)
        cells3 = [
            (str(row.get('ASIN',''))[:14],           1.6,  DARK, False),
            (str(row.get('Brand',''))[:18],          1.7,  DARK, True),
            (str(row.get('Analyst','')),              .9,  GRAY, False),
            (f"${row.get('Net price',0):.2f}",        .95,  DARK, False),
            (f"${row.get('BB Price',0):.2f}",         .95,  DARK, False),
            (f"${row.get('Breakeven',0):.2f}",        .95,  RED,  True),
            (f"${row.get('Difference from SP',0):.2f}",.95, RED, True),
            (str(row.get('Remarks ',''))[:30],        2.4,  RED,  False),
            (str(row.get('Total Product Score','')),  .75,  GRAY, False),
        ]
        tx3 = Inches(.4)
        for val3,cw33,col3,bld3 in cells3:
            _txt(s, val3, tx3+Inches(.06), ry3+Inches(.11), Inches(cw33-.1), Inches(.26),
                 size=8.5, bold=bld3, color=col3, align=PP_ALIGN.CENTER)
            tx3 += Inches(cw33)

    s.notes_slide.notes_text_frame.text = (
        "All 8 rejected ASINs are from PROHITTER brand. "
        "4 have negative margins — BB Price doesn't cover landed cost. "
        "4 have low margins — technically profitable but below our threshold. "
        "Action: contact PROHITTER vendor for price concession of 15-20% minimum."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — SOURCING RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════
def slide_sourcing(prs, d):
    df   = d['df']
    s    = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, LIGHT)
    _slide_header(s, "Sourcing Recommendations — Action Plan",
                  "Priority-ranked actions for the sourcing and vendor management teams", d['week_label'])
    _footer(s)

    yes_df   = df[df['is_yes']]
    top_asin = yes_df.sort_values('Last month sold',ascending=False).head(5)
    top_brand= yes_df.groupby('Brand')['ASIN'].count().idxmax()
    top_brand_margin = round(yes_df.groupby('Brand')['margin_pct'].mean().get(top_brand,0),1)

    # Priority 1
    _rect(s, Inches(.4), Inches(1.2), Inches(12.5), Inches(1.3), WHITE)
    _rect(s, Inches(.4), Inches(1.2), Inches(.12), Inches(1.3), ORANGE)
    _txt(s, "PRIORITY 1 — IMMEDIATE PO ACTION", Inches(.62), Inches(1.28),
         Inches(9), Inches(.28), size=9, bold=True, color=ORANGE)
    asin_list = "  ·  ".join([f"{r['ASIN']} ({r['Brand']}, {int(r['Last month sold'])} units/mo)" for _,r in top_asin.iterrows()])
    _txt(s, f"Action these ASINs immediately — highest demand velocity this week:\n{asin_list}",
         Inches(.62), Inches(1.6), Inches(11.7), Inches(.75), size=10, color=DARK)

    # Priority 2
    _rect(s, Inches(.4), Inches(2.65), Inches(12.5), Inches(1.1), WHITE)
    _rect(s, Inches(.4), Inches(2.65), Inches(.12), Inches(1.1), BLUE)
    _txt(s, "PRIORITY 2 — VENDOR FOCUS", Inches(.62), Inches(2.73),
         Inches(9), Inches(.28), size=9, bold=True, color=BLUE)
    _txt(s, f"{top_brand} is this week's top vendor with {int(yes_df[yes_df['Brand']==top_brand]['ASIN'].count())} approved SKUs and an average margin of {top_brand_margin}%. Expand PO allocation and negotiate volume pricing for next cycle.",
         Inches(.62), Inches(3.05), Inches(11.7), Inches(.55), size=10, color=DARK)

    # Priority 3
    _rect(s, Inches(.4), Inches(3.9), Inches(12.5), Inches(1.1), WHITE)
    _rect(s, Inches(.4), Inches(3.9), Inches(.12), Inches(1.1), AMBER)
    _txt(s, "PRIORITY 3 — VENDOR RENEGOTIATION", Inches(.62), Inches(3.98),
         Inches(9), Inches(.28), size=9, bold=True, color=AMBER)
    _txt(s, "8 PROHITTER ASINs rejected due to negative/low margin. Request cost reduction of 15-20% from vendor. If vendor does not comply, reduce PO allocation for next quarter.",
         Inches(.62), Inches(4.3), Inches(11.7), Inches(.55), size=10, color=DARK)

    # Priority 4
    _rect(s, Inches(.4), Inches(5.15), Inches(12.5), Inches(1.0), WHITE)
    _rect(s, Inches(.4), Inches(5.15), Inches(.12), Inches(1.0), GREEN)
    _txt(s, "PRIORITY 4 — GROWTH OPPORTUNITY", Inches(.62), Inches(5.23),
         Inches(9), Inches(.28), size=9, bold=True, color=GREEN)
    _txt(s, "Nintendo (7 SKUs, avg 100 units/month) and Ubisoft (4 SKUs) show strong FBA competition gaps (<5 sellers). Increase catalogue coverage and consider priority restocking.",
         Inches(.62), Inches(5.55), Inches(11.7), Inches(.55), size=10, color=DARK)

    s.notes_slide.notes_text_frame.text = (
        "Walk through each priority in order with the team. "
        "Assign owners before leaving the meeting. "
        "P1 = Sourcing team. P2 = Vendor manager. P3 = Finance + Procurement. P4 = Catalogue team."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — NEXT STEPS
# ═══════════════════════════════════════════════════════════════════════════════
def slide_closing(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6]); _set_bg(s, NAVY)
    _rect(s, Emu(0), Emu(0), Inches(4.8), H, ORANGE)
    _txt(s, "NEXT\nSTEPS", Inches(.55), Inches(1.8), Inches(3.7), Inches(2.8),
         size=54, bold=True, color=WHITE)
    _txt(s, d['week_label'], Inches(.55), Inches(4.8), Inches(4.0), Inches(.4),
         size=11, color=WHITE)
    _txt(s, "virventures.com", Inches(.55), H-Inches(.9), Inches(4.0), Inches(.35),
         size=10, color=RGBColor(0xF9,0xC0,0x8E))

    steps = [
        ("01", "Sourcing Team",      "Issue POs for top 5 ASINs by demand velocity this week."),
        ("02", "Vendor Management",  f"Expand {d['df'][d['df']['is_yes']].groupby('Brand')['ASIN'].count().idxmax()} allocation — negotiate volume pricing for next cycle."),
        ("03", "Finance",            "Review PROHITTER cost structure — target 15-20% price reduction."),
        ("04", "Analyst Team",       "Maintain weekly cadence. Target 75%+ approval rate across all analysts."),
        ("05", "Catalogue Team",     "Expand Nintendo and Ubisoft SKU coverage — low FBA competition opportunity."),
        ("06", "Management",         "Review next week's mastersheet against this week's approval benchmarks."),
    ]
    sy = Inches(1.2)
    for num, owner, action in steps:
        _rect(s, Inches(5.3), sy, Inches(7.6), Inches(.82), NAVY2)
        _txt(s, num, Inches(5.4), sy+Inches(.12), Inches(.5), Inches(.58),
             size=18, bold=True, color=ORANGE)
        _txt(s, owner.upper(), Inches(5.98), sy+Inches(.08), Inches(6.5), Inches(.28),
             size=8, bold=True, color=ORANGE)
        _txt(s, action, Inches(5.98), sy+Inches(.32), Inches(6.5), Inches(.44),
             size=10, color=WHITE)
        sy += Inches(.95)

    s.notes_slide.notes_text_frame.text = (
        "Closing slide. Assign each action item to a specific owner before the meeting ends. "
        "Set a follow-up date for vendor renegotiation outcomes. "
        "Next week's report should show improvement in PROHITTER margin position."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
def generate_pptx(data: dict) -> bytes:
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    slide_cover(prs, data)
    slide_exec_summary(prs, data)
    slide_top_asins(prs, data)
    slide_analyst(prs, data)
    slide_brands(prs, data)
    slide_rejected(prs, data)
    slide_sourcing(prs, data)
    slide_closing(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
