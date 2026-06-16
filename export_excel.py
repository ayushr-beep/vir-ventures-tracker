"""
Vir Ventures — Boardroom Excel Workbook Generator
5 premium sheets: Executive Dashboard, ASIN Intelligence,
Analyst Scorecards, Brand Intelligence, Rejection Analysis.
"""
import io, numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from datetime import datetime

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY_H   = "0F1629"; NAVY2_H  = "1E2A45"; ORANGE_H = "E8611A"
ORANGE2H = "F4874B"; WHITE_H  = "FFFFFF"; LIGHT_H  = "F8F9FB"
LGRAY_H  = "E5E7EB"; GRAY_H   = "6B7280"; GREEN_H  = "16A34A"
RED_H    = "DC2626"; AMBER_H  = "D97706"; BLUE_H   = "2563EB"
DARK_H   = "111827"; GGREEN_H = "DCFCE7"; LRED_H   = "FEE2E2"
LAMBER_H = "FEF9C3"; LORANGE  = "FFF0E8"

def nf(h): return PatternFill("solid", fgColor=h)
def bw(sz=11,c=WHITE_H,b=True): return Font(name="Calibri",bold=b,size=sz,color=c)
def bd(sz=11,c=DARK_H,b=False): return Font(name="Calibri",bold=b,size=sz,color=c)
def bo(sz=11): return Font(name="Calibri",bold=True,size=sz,color=ORANGE_H)
def bg(sz=11): return Font(name="Calibri",bold=True,size=sz,color=GREEN_H)
def br(sz=11): return Font(name="Calibri",bold=True,size=sz,color=RED_H)
def ba(sz=11): return Font(name="Calibri",bold=True,size=sz,color=AMBER_H)

def cen(): return Alignment(horizontal="center",vertical="center",wrap_text=True)
def lft(): return Alignment(horizontal="left",  vertical="center",wrap_text=True)
def rgt(): return Alignment(horizontal="right", vertical="center")

def thin(c=LGRAY_H):
    s = Side(style="thin",color=c)
    return Border(left=s,right=s,top=s,bottom=s)

def thick_bottom():
    s = Side(style="medium",color=ORANGE_H)
    return Border(bottom=s)

def rate_style(rate):
    if rate>=80: return nf(GGREEN_H), Font(name="Calibri",bold=True,size=11,color=GREEN_H)
    if rate>=50: return nf(LAMBER_H), Font(name="Calibri",bold=True,size=11,color=AMBER_H)
    return nf(LRED_H), Font(name="Calibri",bold=True,size=11,color=RED_H)

def scw(ws, widths):
    for col,w in widths.items(): ws.column_dimensions[col].width = w

def header_row(ws, row, cols, start=1):
    for i,(label,w) in enumerate(cols):
        c = ws.cell(row=row,column=start+i,value=label)
        c.font = bw(9); c.fill = nf(NAVY2_H)
        c.alignment = cen(); c.border = thin()
        ws.column_dimensions[get_column_letter(start+i)].width = w

def banner(ws, row, text, merge_to, height=28):
    ws.merge_cells(f"A{row}:{get_column_letter(merge_to)}{row}")
    c = ws.cell(row=row,column=1,value=text)
    c.font = Font(name="Calibri",bold=True,size=14,color=WHITE_H)
    c.fill = nf(NAVY_H); c.alignment = cen()
    ws.row_dimensions[row].height = height

def sub_banner(ws, row, text, merge_to, height=20):
    ws.merge_cells(f"A{row}:{get_column_letter(merge_to)}{row}")
    c = ws.cell(row=row,column=1,value=text)
    c.font = Font(name="Calibri",bold=True,size=10,color=ORANGE_H)
    c.fill = nf(NAVY2_H); c.alignment = lft()
    ws.row_dimensions[row].height = height

def sg(ws): ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — EXECUTIVE DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def sheet_executive(wb, d):
    ws = wb.create_sheet("📊 Executive Dashboard"); sg(ws)
    df = d['df']

    banner(ws, 1, f"VIR VENTURES — BOARDROOM ANALYST DASHBOARD  ·  {d['week_label'].upper()}", 12, 36)
    ws.merge_cells("A2:L2")
    c2 = ws.cell(row=2,column=1, value=f"Generated {datetime.now().strftime('%d %b %Y %H:%M')}  ·  Confidential — Internal Management Use Only  ·  Report Date: {d['report_date']}")
    c2.font = Font(name="Calibri",size=9,color=ORANGE_H); c2.fill = nf(NAVY2_H); c2.alignment = cen()
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 8

    # KPI cards
    kpi_data = [
        ("A", "CATALOGUE SIZE",   d['total'],       "Total SKUs Reviewed",     ORANGE_H, DARK_H),
        ("D", "APPROVED",         d['approved'],    "Yes + Yes, Need Discount", GREEN_H,  GREEN_H),
        ("G", "REJECTED",         d['rejected'],    "Negative / Low Margin",    RED_H,    RED_H),
        ("J", "APPROVAL RATE",    f"{d['rate']}%",  f"of {d['total']} SKUs",   BLUE_H,   BLUE_H),
    ]
    for col_letter, label, val, sub, accent, vcol in kpi_data:
        col_n = ord(col_letter)-64
        ws.merge_cells(start_row=4,end_row=4,start_column=col_n,end_column=col_n+1)
        ws.merge_cells(start_row=5,end_row=6,start_column=col_n,end_column=col_n+1)
        ws.merge_cells(start_row=7,end_row=7,start_column=col_n,end_column=col_n+1)
        lc = ws.cell(row=4,column=col_n,value=label)
        lc.font = Font(name="Calibri",bold=True,size=8,color=GRAY_H); lc.fill = nf(WHITE_H); lc.alignment = lft()
        vc = ws.cell(row=5,column=col_n,value=val)
        vc.font = Font(name="Calibri",bold=True,size=24,color=vcol); vc.fill = nf(WHITE_H); vc.alignment = cen()
        sc = ws.cell(row=7,column=col_n,value=sub)
        sc.font = Font(name="Calibri",size=8,color=GRAY_H); sc.fill = nf(WHITE_H); sc.alignment = lft()
        ws.row_dimensions[4].height = 14; ws.row_dimensions[5].height = 26
        ws.row_dimensions[6].height = 10; ws.row_dimensions[7].height = 14

    ws.row_dimensions[8].height = 10

    # Top Brand & Analyst highlights
    yes_df = df[df['is_yes']]
    top_analyst = df.groupby('Analyst')['is_yes'].mean().idxmax()
    top_rate    = round(df.groupby('Analyst')['is_yes'].mean()[top_analyst]*100,1)
    top_brand   = yes_df.groupby('Brand')['ASIN'].count().idxmax() if not yes_df.empty else "—"
    top_brand_n = int(yes_df.groupby('Brand')['ASIN'].count().max()) if not yes_df.empty else 0
    top_asin_r  = yes_df.sort_values('Last month sold',ascending=False).iloc[0] if not yes_df.empty else None
    avg_margin  = round(yes_df['margin_pct'].mean(),1) if not yes_df.empty else 0

    highlights = [
        ("TOP ANALYST",      top_analyst,                     f"{top_rate}% Approval Rate"),
        ("TOP BRAND",        top_brand,                       f"{top_brand_n} Approved SKUs"),
        ("TOP ASIN",         top_asin_r['ASIN'] if top_asin_r is not None else "—",
                             f"{int(top_asin_r['Last month sold'])} units/mo" if top_asin_r is not None else "—"),
        ("AVG MARGIN (APP)", f"{avg_margin}%",                "On approved SKUs"),
    ]
    hl_start = 9
    for i,(label,val,sub) in enumerate(highlights):
        col_n = 1+i*3
        ws.merge_cells(start_row=hl_start,end_row=hl_start,start_column=col_n,end_column=col_n+1)
        ws.merge_cells(start_row=hl_start+1,end_row=hl_start+2,start_column=col_n,end_column=col_n+1)
        ws.merge_cells(start_row=hl_start+3,end_row=hl_start+3,start_column=col_n,end_column=col_n+1)
        lc = ws.cell(row=hl_start,column=col_n,value=label)
        lc.font = Font(name="Calibri",bold=True,size=8,color=ORANGE_H); lc.fill = nf(LORANGE); lc.alignment = lft()
        vc = ws.cell(row=hl_start+1,column=col_n,value=val)
        vc.font = Font(name="Calibri",bold=True,size=16,color=DARK_H); vc.fill = nf(WHITE_H); vc.alignment = cen()
        sc = ws.cell(row=hl_start+3,column=col_n,value=sub)
        sc.font = Font(name="Calibri",size=8,color=GRAY_H); sc.fill = nf(WHITE_H); sc.alignment = lft()
        ws.row_dimensions[hl_start].height   = 14
        ws.row_dimensions[hl_start+1].height = 22
        ws.row_dimensions[hl_start+2].height = 8
        ws.row_dimensions[hl_start+3].height = 14

    ws.row_dimensions[hl_start+4].height = 10

    # Analyst summary table
    r = hl_start+5
    sub_banner(ws, r, "ANALYST PERFORMANCE SUMMARY", 12, 20)

    analyst_stats = df.groupby('Analyst').agg(
        Total=('ASIN','count'), Approved=('is_yes','sum'),
        Rejected=('is_yes',lambda x:(~x).sum()),
        Rate=('is_yes',lambda x:round(x.mean()*100,1)),
        AvgMargin=('margin_pct','mean'), TotalUnits=('Last month sold','sum'),
        AvgRank=('Rank','mean'), AvgNetPrice=('Net price','mean'),
        AvgBBPrice=('BB Price','mean'), AvgDemand=('Total Demand Score','mean'),
        AvgComp=('Total Competition Score','mean'), AvgMarginScore=('Total Margin Score','mean'),
    ).reset_index().round(1)

    r += 1
    header_row(ws, r, [
        ("Analyst",14),("Total Audited",14),("Approved",12),("Rejected",10),
        ("Approval Rate",13),("Avg Margin %",13),("Total Units/Mo",14),("Avg BSR Rank",13),
        ("Avg Net Price",13),("Avg BB Price",13),("Avg Demand Sc.",13),
        ("Avg Comp. Sc.",13)], 1)
    ws.row_dimensions[r].height = 22

    for ri,(_,row) in enumerate(analyst_stats.iterrows()):
        r+=1
        rate = row['Rate']
        rf,rfont = rate_style(rate)
        vals = [row['Analyst'],int(row['Total']),int(row['Approved']),int(row['Rejected']),
                rate, row['AvgMargin'], int(row['TotalUnits']), round(row['AvgRank']),
                row['AvgNetPrice'], row['AvgBBPrice'], row['AvgDemand'], row['AvgComp']]
        row_bg = nf(WHITE_H) if ri%2==0 else nf(LIGHT_H)
        for ci,val in enumerate(vals,1):
            c = ws.cell(row=r,column=ci,value=val)
            c.fill = rf if ci==5 else row_bg
            c.font = rfont if ci==5 else bd(10,b=(ci==1 and ri==0))
            c.alignment = cen(); c.border = thin()
        ws.row_dimensions[r].height = 20

    r+=2
    # Brand summary table
    sub_banner(ws, r, "BRAND PERFORMANCE SUMMARY", 7, 20)
    brand_stats = df[df['is_yes']].groupby('Brand').agg(
        Approved=('ASIN','count'), TotalUnits=('Last month sold','sum'),
        AvgMarginPct=('margin_pct','mean'), AvgRank=('Rank','mean'),
        AvgBBPrice=('BB Price','mean'),
    ).reset_index().sort_values('Approved',ascending=False).round(1)

    r+=1
    header_row(ws, r, [("Brand",20),("Approved SKUs",14),("Total Units/Mo",14),
                        ("Avg Margin %",13),("Avg BSR Rank",13),("Avg BB Price",13),
                        ("Priority",10)], 1)
    ws.row_dimensions[r].height = 22
    for ri,(_,row) in enumerate(brand_stats.iterrows()):
        r+=1
        priority = "🔥 IMMEDIATE" if ri==0 else ("★ HIGH" if ri<3 else "PIPELINE")
        vals = [row['Brand'],int(row['Approved']),int(row['TotalUnits']),
                row['AvgMarginPct'], round(row['AvgRank']), row['AvgBBPrice'], priority]
        row_bg = nf("FFF7ED") if ri==0 else (nf(WHITE_H) if ri%2==0 else nf(LIGHT_H))
        for ci,val in enumerate(vals,1):
            c = ws.cell(row=r,column=ci,value=val)
            c.fill = row_bg; c.border = thin()
            c.font = bd(10,b=(ri==0))
            if ci==4: c.font = bo(10)
            c.alignment = cen()
        ws.row_dimensions[r].height = 20

    scw(ws,{"A":14,"B":14,"C":12,"D":10,"E":13,"F":13,"G":14,"H":13,"I":13,"J":13,"K":13,"L":13})
    ws.freeze_panes = "A9"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — ASIN INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
def sheet_asins(wb, d):
    ws = wb.create_sheet("📦 ASIN Intelligence"); sg(ws)
    df = d['df']
    yes_df = df[df['is_yes']].sort_values('Last month sold',ascending=False).copy()

    banner(ws, 1, "ASIN INTELLIGENCE — APPROVED SKUs RANKED BY DEMAND VELOCITY", 14, 32)
    ws.row_dimensions[2].height = 8

    r = 3
    sub_banner(ws, r, f"  APPROVED ASINs  ·  {len(yes_df)} SKUs  ·  All values from Week 1 Mastersheet", 14, 20)

    r+=1
    cols = [
        ("#",4),("ASIN",14),("Brand",20),("Analyst",11),("Recommended",18),
        ("Units/Month",13),("Net Price ($)",13),("BB Price ($)",13),
        ("Margin Gap ($)",14),("Margin %",12),("BSR Rank",11),
        ("Product Score",13),("FBA Sellers",12),("Buy Box Holder",18),("Remarks",28)
    ]
    header_row(ws, r, cols, 1)
    ws.row_dimensions[r].height = 22

    for ri,(_,row) in enumerate(yes_df.iterrows()):
        r+=1
        margin_p = row.get('margin_pct',0)
        remark = str(row.get('Remarks ',''))

        vals = [ri+1, row['ASIN'], row['Brand'], row['Analyst'], row['Recommended'],
                int(row.get('Last month sold',0)), row.get('Net price',0), row.get('BB Price',0),
                row.get('margin_gap',0), margin_p, int(row.get('Rank',0)),
                row.get('Total Product Score',''), int(row.get('Number of FBA vendors',0)),
                row.get('BuyBoxSellerName',''), remark]

        # Row bg by demand tier
        units = int(row.get('Last month sold',0))
        if units>=200: bg_h = "F0FDF4"
        elif units>=100: bg_h = "FFF7ED"
        elif units>=50: bg_h = LIGHT_H
        else: bg_h = WHITE_H

        for ci,val in enumerate(vals,1):
            c = ws.cell(row=r,column=ci,value=val)
            c.fill = nf(bg_h); c.border = thin(); c.alignment = cen()
            c.font = bd(9,b=(ri==0))
            # Special formatting
            if ci==6 and units>0: c.font = Font(name="Calibri",bold=True,size=9,color=GREEN_H)
            if ci==10: c.font = Font(name="Calibri",bold=True,size=9,color=ORANGE_H)
            if ci==9 and val and float(val)<0: c.font = Font(name="Calibri",bold=True,size=9,color=RED_H)
        ws.row_dimensions[r].height = 18

    # Trophy labels for top 3
    for rx in range(4, 7):
        c = ws.cell(row=rx,column=1)
        if rx==4: c.value = "🥇"
        elif rx==5: c.value = "🥈"
        elif rx==6: c.value = "🥉"

    r+=2
    # Summary stats
    sub_banner(ws, r, "  ASIN PERFORMANCE SUMMARY STATISTICS", 8, 18)
    r+=1
    stats_rows = [
        ("Total Approved ASINs",    len(yes_df)),
        ("Total Units/Month",       int(yes_df['Last month sold'].sum())),
        ("Avg Units/Month",         round(yes_df['Last month sold'].mean(),1)),
        ("Avg Buy Box Margin %",    round(yes_df['margin_pct'].mean(),1)),
        ("Avg BSR Rank",            round(yes_df['Rank'].mean())),
        ("Highest Demand ASIN",     yes_df.iloc[0]['ASIN'] if not yes_df.empty else "—"),
        ("Best Margin ASIN",        yes_df.sort_values('margin_pct',ascending=False).iloc[0]['ASIN'] if not yes_df.empty else "—"),
        ("Amazon in Buy Box (%)",   round((yes_df['BuyBoxSellerName']=='Amazon.com').mean()*100,1)),
    ]
    for label,val in stats_rows:
        c1 = ws.cell(row=r,column=1,value=label); c1.font = bd(10,b=True); c1.fill = nf(LIGHT_H); c1.alignment = lft()
        c2 = ws.cell(row=r,column=2,value=val);   c2.font = bo(10);       c2.fill = nf(LORANGE);  c2.alignment = cen()
        ws.row_dimensions[r].height = 20; r+=1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(15)}{4+len(yes_df)}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — ANALYST SCORECARDS
# ══════════════════════════════════════════════════════════════════════════════
def sheet_analysts(wb, d):
    ws = wb.create_sheet("👥 Analyst Scorecards"); sg(ws)
    df = d['df']

    banner(ws, 1, "ANALYST PERFORMANCE SCORECARDS — WEEK 1", 10, 32)
    ws.row_dimensions[2].height = 8

    analyst_stats = df.groupby('Analyst').agg(
        Total=('ASIN','count'), Approved=('is_yes','sum'),
        Rejected=('is_yes',lambda x:(~x).sum()),
        Rate=('is_yes',lambda x:round(x.mean()*100,1)),
        AvgMargin=('margin_pct','mean'), TotalUnits=('Last month sold','sum'),
        AvgRank=('Rank','mean'), AvgNetPrice=('Net price','mean'),
        AvgBBPrice=('BB Price','mean'), AvgDemand=('Total Demand Score','mean'),
        AvgComp=('Total Competition Score','mean'), AvgMarginScore=('Total Margin Score','mean'),
    ).reset_index().sort_values('Rate',ascending=False).round(1)

    r = 3
    for _,row in analyst_stats.iterrows():
        rate = row['Rate']
        accent_h = GREEN_H if rate>=80 else (AMBER_H if rate>=50 else RED_H)
        accent_l = GGREEN_H if rate>=80 else (LAMBER_H if rate>=50 else LRED_H)

        # Analyst banner
        ws.merge_cells(f"A{r}:J{r}")
        c = ws.cell(row=r,column=1,value=f"  {row['Analyst'].upper()}  ·  {rate}% Approval Rate  ·  {int(row['Total'])} SKUs Audited")
        c.font = Font(name="Calibri",bold=True,size=12,color=WHITE_H)
        c.fill = nf(NAVY2_H); c.alignment = lft()
        ws.row_dimensions[r].height = 24; r+=1

        # Score bar
        ws.merge_cells(f"A{r}:J{r}")
        c = ws.cell(row=r,column=1,value=f"  Approval Rate: {rate}%  |  Approved: {int(row['Approved'])}  |  Rejected: {int(row['Rejected'])}  |  Avg Margin: {row['AvgMargin']:.1f}%  |  Total Units/Mo: {int(row['TotalUnits'])}")
        c.font = Font(name="Calibri",bold=True,size=10,color=accent_h)
        c.fill = nf(accent_l); c.alignment = lft()
        ws.row_dimensions[r].height = 18; r+=1

        # Metrics grid header
        metric_cols = [("Metric",22),("Value",14),("Metric",22),("Value",14)]
        for ci,(label,w) in enumerate(metric_cols,1):
            c = ws.cell(row=r,column=ci,value=label)
            c.font = bw(9); c.fill = nf(NAVY_H); c.alignment = cen()
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[r].height = 18; r+=1

        # Metrics in 2 columns
        metrics_left = [
            ("Total SKUs Audited",      int(row['Total'])),
            ("Approved",                int(row['Approved'])),
            ("Rejected",                int(row['Rejected'])),
            ("Approval Rate",           f"{rate}%"),
            ("Avg Margin % (on app.)",  f"{row['AvgMargin']:.1f}%"),
            ("Total Units/Month",       int(row['TotalUnits'])),
        ]
        metrics_right = [
            ("Avg BSR Rank",            round(row['AvgRank'])),
            ("Avg Net Price ($)",        row['AvgNetPrice']),
            ("Avg BB Price ($)",         row['AvgBBPrice']),
            ("Avg Demand Score",         row['AvgDemand']),
            ("Avg Competition Score",    row['AvgComp']),
            ("Avg Margin Score",         row['AvgMarginScore']),
        ]
        for i,(ml,mr) in enumerate(zip(metrics_left,metrics_right)):
            bg = nf(WHITE_H) if i%2==0 else nf(LIGHT_H)
            for ci,val in [(1,ml[0]),(2,ml[1]),(3,mr[0]),(4,mr[1])]:
                c = ws.cell(row=r,column=ci,value=val)
                c.fill = bg; c.border = thin(); c.alignment = cen()
                if ci in (1,3): c.font = bd(10,b=True)
                elif ci==2: c.font = bo(10)
                else: c.font = bd(10)
            ws.row_dimensions[r].height = 18; r+=1

        # ASIN breakdown for this analyst
        analyst_asins = df[df['Analyst']==row['Analyst']].sort_values('Last month sold',ascending=False)
        r+=1
        sub_banner(ws, r, f"  {row['Analyst']} — ASIN-by-ASIN REVIEW", 10, 18); r+=1
        header_row(ws, r, [("ASIN",14),("Brand",20),("Recommended",18),
                             ("Units/Mo",11),("Net $",10),("BB $",10),
                             ("Margin%",10),("BSR Rank",10),
                             ("Score",9),("Remarks",25)], 1)
        ws.row_dimensions[r].height = 20; r+=1

        for ri2,(_,arow) in enumerate(analyst_asins.iterrows()):
            rec = arow['Recommended']
            bg2 = nf(GGREEN_H) if arow['is_yes'] else nf(LRED_H)
            vals2 = [arow['ASIN'],arow['Brand'],rec,
                     int(arow.get('Last month sold',0)),arow.get('Net price',0),
                     arow.get('BB Price',0),arow.get('margin_pct',0),
                     int(arow.get('Rank',0)),arow.get('Total Product Score',''),
                     str(arow.get('Remarks ',''))]
            for ci,val in enumerate(vals2,1):
                c = ws.cell(row=r,column=ci,value=val)
                c.fill = bg2; c.border = thin(); c.alignment = cen()
                c.font = Font(name="Calibri",size=9,
                              color=GREEN_H if arow['is_yes'] else RED_H if ci==3 else DARK_H)
            ws.row_dimensions[r].height = 17; r+=1

        r+=2  # spacer between analysts

    ws.freeze_panes = "A4"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — BRAND INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
def sheet_brands(wb, d):
    ws = wb.create_sheet("🏷️ Brand Intelligence"); sg(ws)
    df = d['df']

    banner(ws, 1, "BRAND INTELLIGENCE — VENDOR PERFORMANCE & SOURCING PRIORITY", 9, 32)
    ws.row_dimensions[2].height = 8

    brand_stats = df.groupby('Brand').agg(
        Total=('ASIN','count'), Approved=('is_yes','sum'),
        Rejected=('is_yes',lambda x:(~x).sum()),
        Rate=('is_yes',lambda x:round(x.mean()*100,1)),
        TotalUnits=('Last month sold','sum'),
        AvgMarginPct=('margin_pct','mean'), AvgRank=('Rank','mean'),
        AvgNetPrice=('Net price','mean'), AvgBBPrice=('BB Price','mean'),
        AvgFBASellers=('Number of FBA vendors','mean'),
    ).reset_index().sort_values('Approved',ascending=False).round(1)
    brand_stats['Priority'] = ['🔥 IMMEDIATE' if i==0 else ('★ HIGH' if i<3 else 'PIPELINE') for i in range(len(brand_stats))]

    r = 3
    sub_banner(ws, r, f"  BRAND RANKING BY APPROVED SKUs  ·  {len(brand_stats)} Brands", 10, 20); r+=1
    header_row(ws, r, [
        ("Rank",6),("Brand",20),("Total SKUs",12),("Approved",12),("Rejected",10),
        ("Approval %",12),("Total Units/Mo",14),("Avg Margin %",13),
        ("Avg BSR Rank",13),("Avg FBA Sellers",14),("Priority",14)], 1)
    ws.row_dimensions[r].height = 22; r+=1

    for ri,(_,row) in enumerate(brand_stats.iterrows()):
        bg = nf("FFF7ED") if ri==0 else (nf("FFFBEB") if ri<3 else (nf(WHITE_H) if ri%2==0 else nf(LIGHT_H)))
        vals = [ri+1, row['Brand'], int(row['Total']), int(row['Approved']), int(row['Rejected']),
                row['Rate'], int(row['TotalUnits']), row['AvgMarginPct'],
                round(row['AvgRank']), row['AvgFBASellers'], row['Priority']]
        for ci,val in enumerate(vals,1):
            c = ws.cell(row=r,column=ci,value=val)
            c.fill = bg; c.border = thin(); c.alignment = cen()
            c.font = bd(10,b=(ri==0))
            if ci==8: c.font = Font(name="Calibri",bold=True,size=10,color=ORANGE_H)
            if ci==6:
                rf2,rf2font = rate_style(row['Rate'])
                c.fill = rf2; c.font = rf2font
        ws.row_dimensions[r].height = 20; r+=1

    # Bar chart
    r+=2
    chart_data_start = 5
    chart_data_end   = 5 + len(brand_stats) - 1
    chart = BarChart(); chart.type = "bar"
    chart.title = "Approved SKUs by Brand"; chart.style = 10
    chart.y_axis.title = "Brand"; chart.x_axis.title = "Approved SKUs"
    chart.width = 18; chart.height = 12
    data_ref = Reference(ws, min_col=4, min_row=chart_data_start, max_row=chart_data_end)
    cats_ref = Reference(ws, min_col=2, min_row=chart_data_start, max_row=chart_data_end)
    chart.add_data(data_ref)
    chart.set_categories(cats_ref)
    chart.series[0].graphicalProperties.solidFill = ORANGE_H
    ws.add_chart(chart, f"A{r}")

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(11)}{4+len(brand_stats)}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — REJECTION ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def sheet_rejected(wb, d):
    ws = wb.create_sheet("⚠️ Rejection Analysis"); sg(ws)
    df = d['df']
    rej = df[~df['is_yes']].copy()

    banner(ws, 1, "REJECTION ANALYSIS — ROOT CAUSE & VENDOR ACTION PLAN", 10, 32)
    ws.row_dimensions[2].height = 8

    r = 3
    # Summary
    neg = int(rej['Remarks '].str.contains('Negative',na=False).sum())
    low = int(rej['Remarks '].str.contains('Low',na=False).sum())
    oth = len(rej) - neg - low

    summary = [
        ("Total Rejected",    len(rej),  "All flagged this week"),
        ("Negative Margin",   neg,       "BB Price < Breakeven — immediate action"),
        ("Low Margin",        low,       "Marginal profitability — negotiate"),
        ("Other Reasons",     oth,       "Requires case-by-case review"),
    ]
    for col_n,(label,val,sub) in enumerate(summary,1):
        col_n_actual = (col_n-1)*3+1
        ws.merge_cells(start_row=r,end_row=r,start_column=col_n_actual,end_column=col_n_actual+1)
        ws.merge_cells(start_row=r+1,end_row=r+2,start_column=col_n_actual,end_column=col_n_actual+1)
        ws.merge_cells(start_row=r+3,end_row=r+3,start_column=col_n_actual,end_column=col_n_actual+1)
        lc = ws.cell(row=r,column=col_n_actual,value=label)
        lc.font = Font(name="Calibri",bold=True,size=8,color=RED_H); lc.fill = nf(LRED_H); lc.alignment = lft()
        vc = ws.cell(row=r+1,column=col_n_actual,value=val)
        vc.font = Font(name="Calibri",bold=True,size=22,color=RED_H); vc.fill = nf(WHITE_H); vc.alignment = cen()
        sc = ws.cell(row=r+3,column=col_n_actual,value=sub)
        sc.font = Font(name="Calibri",size=8,color=GRAY_H); sc.fill = nf(WHITE_H); sc.alignment = lft()
        ws.row_dimensions[r].height   = 14; ws.row_dimensions[r+1].height = 22
        ws.row_dimensions[r+2].height = 8;  ws.row_dimensions[r+3].height = 14

    r+=5
    ws.row_dimensions[r].height = 8; r+=1
    sub_banner(ws, r, "  REJECTED ASIN DETAIL — FULL BREAKDOWN", 10, 20); r+=1
    header_row(ws, r, [("ASIN",14),("Brand",16),("Analyst",11),
                        ("Net Price",11),("BB Price",11),("Breakeven",11),
                        ("Diff from SP",12),("Margin Gap",11),("Product Score",13),
                        ("Reason",18)], 1)
    ws.row_dimensions[r].height = 22; r+=1

    for _,row in rej.iterrows():
        gap = row.get('BB Price',0) - row.get('Net price',0)
        vals = [row['ASIN'],row['Brand'],row['Analyst'],
                row.get('Net price',0),row.get('BB Price',0),row.get('Breakeven',0),
                row.get('Difference from SP',0),round(gap,2),
                row.get('Total Product Score',''),str(row.get('Remarks ',''))]
        for ci,val in enumerate(vals,1):
            c = ws.cell(row=r,column=ci,value=val)
            c.fill = nf(LRED_H) if ci in (4,5,6,7) else nf(WHITE_H)
            c.border = thin(); c.alignment = cen()
            c.font = Font(name="Calibri",size=9,
                          color=RED_H if ci in (4,5,6,7) else DARK_H,
                          bold=(ci==2))
        ws.row_dimensions[r].height = 18; r+=1

    r+=2
    # Action plan box
    sub_banner(ws, r, "  MANAGEMENT ACTION PLAN FOR REJECTED SKUs", 10, 20); r+=1
    actions = [
        ("1", "IMMEDIATE", f"{neg} ASINs have Negative Margin. BB Price is below our landed cost. Contact PROHITTER vendor and request minimum 15-20% cost reduction. Do not order until resolved."),
        ("2", "THIS WEEK",  f"{low} ASINs have Low Margin. Evaluate if volume purchase can improve unit economics. Request promotional pricing or extended payment terms from vendor."),
        ("3", "NEXT CYCLE", f"All {len(rej)} rejected ASINs are from PROHITTER (Vendor: PROH). Review vendor relationship and set minimum margin threshold for future submissions."),
    ]
    for num, urgency, action in actions:
        ws.merge_cells(f"A{r}:J{r}")
        c = ws.cell(row=r,column=1,value=f"  [{num}] {urgency}:  {action}")
        c.font = Font(name="Calibri",size=10,color=DARK_H if num!='1' else RED_H)
        c.fill = nf(LRED_H) if num=='1' else nf(LAMBER_H) if num=='2' else nf(LIGHT_H)
        c.alignment = lft(); ws.row_dimensions[r].height = 24; r+=1

    ws.freeze_panes = "A9"

# ══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY
# ══════════════════════════════════════════════════════════════════════════════
def generate_excel(data: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    sheet_executive(wb, data)
    sheet_asins(wb, data)
    sheet_analysts(wb, data)
    sheet_brands(wb, data)
    sheet_rejected(wb, data)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
