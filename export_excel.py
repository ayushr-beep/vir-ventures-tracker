"""
Premium Excel report generator for Vir Ventures Weekly Analyst Tracker.
Produces a multi-sheet executive workbook with full formatting.
"""
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from datetime import datetime

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY_HEX   = "1A1F36"
ORANGE_HEX = "E8611A"
ORANGE2    = "F4874B"
WHITE_HEX  = "FFFFFF"
LIGHT_HEX  = "F8F9FB"
GRAY_HEX   = "6B7280"
LGRAY_HEX  = "E5E7EB"
GREEN_HEX  = "16A34A"
RED_HEX    = "DC2626"
DARK_HEX   = "111827"
AMBER_HEX  = "D97706"
HEADER_BG  = "252B4A"

def navy_fill():   return PatternFill("solid", fgColor=NAVY_HEX)
def orange_fill(): return PatternFill("solid", fgColor=ORANGE_HEX)
def light_fill():  return PatternFill("solid", fgColor=LIGHT_HEX)
def white_fill():  return PatternFill("solid", fgColor=WHITE_HEX)
def lgray_fill():  return PatternFill("solid", fgColor=LGRAY_HEX)
def header_fill(): return PatternFill("solid", fgColor=HEADER_BG)

def bold_white(size=11):  return Font(name="Calibri", bold=True, color=WHITE_HEX, size=size)
def bold_dark(size=11):   return Font(name="Calibri", bold=True, color=DARK_HEX, size=size)
def bold_orange(size=11): return Font(name="Calibri", bold=True, color=ORANGE_HEX, size=size)
def reg(size=11):         return Font(name="Calibri", size=size, color=DARK_HEX)
def gray_font(size=10):   return Font(name="Calibri", size=size, color=GRAY_HEX)

def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left():   return Alignment(horizontal="left",   vertical="center", wrap_text=True)
def right():  return Alignment(horizontal="right",  vertical="center")

def thin_border():
    s = Side(style="thin", color=LGRAY_HEX)
    return Border(left=s, right=s, top=s, bottom=s)

def rate_fill(rate):
    if rate >= 70: return PatternFill("solid", fgColor="DCFCE7")
    if rate >= 40: return PatternFill("solid", fgColor="FEF9C3")
    return PatternFill("solid", fgColor="FEE2E2")

def rate_font(rate):
    c = GREEN_HEX if rate >= 70 else (AMBER_HEX if rate >= 40 else RED_HEX)
    return Font(name="Calibri", bold=True, size=11, color=c)

def set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def write_header_row(ws, row, cols, start_col=1):
    for i, (label, width) in enumerate(cols):
        cell = ws.cell(row=row, column=start_col + i, value=label)
        cell.font = bold_white(10)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = center()
        cell.border = thin_border()
        col_letter = get_column_letter(start_col + i)
        ws.column_dimensions[col_letter].width = width

# ── Sheet 1: Dashboard ────────────────────────────────────────────────────────
def build_dashboard(wb, data):
    ws = wb.create_sheet("📊 Dashboard")
    ws.sheet_view.showGridLines = False

    kpis = data["kpis"]
    a_df = data["analyst_df"]
    brand_df = data["brand_df"]
    bc_col = data["bc_col"]
    week_label = data["week_label"]

    # ── Title block ──
    ws.merge_cells("A1:L1")
    ws["A1"] = "VIR VENTURES — WEEKLY ANALYST PERFORMANCE REPORT"
    ws["A1"].font = Font(name="Calibri", bold=True, size=16, color=WHITE_HEX)
    ws["A1"].fill = navy_fill()
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:L2")
    ws["A2"] = f"{week_label}  ·  Generated {datetime.now().strftime('%d %b %Y %H:%M')}  ·  CONFIDENTIAL"
    ws["A2"].font = Font(name="Calibri", size=10, color=ORANGE_HEX)
    ws["A2"].fill = PatternFill("solid", fgColor="252B4A")
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10

    # ── KPI row ──
    ws.merge_cells("A4:A5"); ws["A4"] = "CATALOGUE SIZE"
    ws.merge_cells("D4:D5"); ws["D4"] = "RECOMMENDED"
    ws.merge_cells("G4:G5"); ws["G4"] = "REJECTED"
    ws.merge_cells("J4:J5"); ws["J4"] = "APPROVAL RATE"

    ws.merge_cells("B4:C5"); ws["B4"] = kpis["total"]
    ws.merge_cells("E4:F5"); ws["E4"] = kpis["recommended"]
    ws.merge_cells("H4:I5"); ws["H4"] = kpis["rejected"]
    ws.merge_cells("K4:L5"); ws["K4"] = f"{kpis['rate']}%"

    label_cells = ["A4", "D4", "G4", "J4"]
    value_cells = ["B4", "E4", "H4", "K4"]
    accents = [ORANGE_HEX, GREEN_HEX, RED_HEX, "2563EB"]

    for lc, vc, acc in zip(label_cells, value_cells, accents):
        ws[lc].font = Font(name="Calibri", bold=True, size=9, color=GRAY_HEX)
        ws[lc].fill = white_fill()
        ws[lc].alignment = Alignment(horizontal="left", vertical="top")
        ws[vc].font = Font(name="Calibri", bold=True, size=26, color=acc)
        ws[vc].fill = white_fill()
        ws[vc].alignment = center()
        ws.row_dimensions[4].height = 18
        ws.row_dimensions[5].height = 28

    ws.row_dimensions[6].height = 10

    # ── Analyst table ──
    if not a_df.empty:
        ws.merge_cells("A7:F7")
        ws["A7"] = "ANALYST PERFORMANCE"
        ws["A7"].font = bold_orange(11)
        ws["A7"].fill = light_fill()
        ws["A7"].alignment = left()
        ws.row_dimensions[7].height = 22

        cols = [("Analyst", 18), ("Total Audited", 15), ("Recommended", 15),
                ("Rejected", 13), ("Approval Rate %", 16), ("Signal", 10)]
        write_header_row(ws, 8, cols, 1)
        ws.row_dimensions[8].height = 22

        for ri, (_, row) in enumerate(a_df.iterrows()):
            r = 9 + ri
            rate = row["Approval Rate %"]
            signal = "🟢 High" if rate >= 70 else ("🟡 Mid" if rate >= 40 else "🔴 Low")
            values = [row["Analyst"], int(row["Total Audited"]),
                      int(row["Recommended"]), int(row["Rejected"]), rate, signal]
            row_fill = white_fill() if ri % 2 == 0 else light_fill()
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=r, column=ci, value=val)
                cell.font = bold_dark(10) if ri == 0 else reg(10)
                cell.fill = row_fill
                cell.alignment = center()
                cell.border = thin_border()
                if ci == 5:
                    cell.fill = rate_fill(rate)
                    cell.font = rate_font(rate)
                    cell.number_format = "0.0%"
            ws.row_dimensions[r].height = 20

    # ── Brand table ──
    if not brand_df.empty and bc_col:
        br = 7
        bc = 8
        ws.cell(row=br, column=bc, value="TOP BRANDS BY APPROVED SKUs")
        ws.cell(row=br, column=bc).font = bold_orange(11)
        ws.cell(row=br, column=bc).fill = light_fill()
        ws.cell(row=br, column=bc).alignment = left()
        ws.merge_cells(start_row=br, end_row=br, start_column=bc, end_column=bc + 2)
        ws.row_dimensions[br].height = 22

        brand_cols = [("Brand", 22), ("Approved SKUs", 16), ("Share", 12)]
        write_header_row(ws, br + 1, brand_cols, bc)
        total_approved = brand_df["Approved SKUs"].sum()
        name_col = bc_col if bc_col in brand_df.columns else brand_df.columns[0]
        for ri, (_, row) in enumerate(brand_df.head(12).iterrows()):
            r = br + 2 + ri
            count = int(row["Approved SKUs"])
            share = count / total_approved if total_approved else 0
            row_fill = white_fill() if ri % 2 == 0 else light_fill()
            for ci, (val, col) in enumerate([(row[name_col], bc), (count, bc + 1), (f"{share*100:.1f}%", bc + 2)]):
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = bold_dark(10) if ri == 0 else reg(10)
                cell.fill = row_fill
                cell.alignment = center()
                cell.border = thin_border()
            ws.row_dimensions[r].height = 20

    # Freeze top rows
    ws.freeze_panes = "A7"


# ── Sheet 2: Analyst Detail ───────────────────────────────────────────────────
def build_analyst_sheet(wb, data):
    ws = wb.create_sheet("👥 Analyst Detail")
    ws.sheet_view.showGridLines = False

    a_df = data["analyst_df"]
    df_full = data["df_full"]
    col_map = data["col_map"]

    ws.merge_cells("A1:I1")
    ws["A1"] = "ANALYST PERFORMANCE — DETAILED VIEW"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE_HEX)
    ws["A1"].fill = navy_fill()
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.row_dimensions[2].height = 8

    if not a_df.empty:
        summary_cols = [("Analyst", 18), ("Total Audited", 15), ("Recommended", 15),
                        ("Rejected", 13), ("Approval Rate %", 16), ("Signal", 10)]
        write_header_row(ws, 3, summary_cols, 1)
        ws.row_dimensions[3].height = 22

        for ri, (_, row) in enumerate(a_df.iterrows()):
            r = 4 + ri
            rate = row["Approval Rate %"]
            signal = "🟢 High" if rate >= 70 else ("🟡 Mid" if rate >= 40 else "🔴 Low")
            vals = [row["Analyst"], int(row["Total Audited"]), int(row["Recommended"]),
                    int(row["Rejected"]), rate, signal]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(row=r, column=ci, value=val)
                cell.font = reg(10)
                cell.fill = white_fill() if ri % 2 == 0 else light_fill()
                cell.alignment = center()
                cell.border = thin_border()
                if ci == 5:
                    cell.fill = rate_fill(rate)
                    cell.font = rate_font(rate)
            ws.row_dimensions[r].height = 20

    # Analyst-level ASIN breakdown
    ac = col_map.get("analyst")
    rc = col_map.get("recommended")
    bc = col_map.get("brand")

    if ac and rc and bc and ac in df_full.columns:
        sep_row = len(a_df) + 6
        ws.cell(row=sep_row, column=1, value="ASIN BREAKDOWN BY ANALYST")
        ws.cell(row=sep_row, column=1).font = bold_orange(11)
        ws.cell(row=sep_row, column=1).fill = light_fill()
        ws.merge_cells(start_row=sep_row, end_row=sep_row, start_column=1, end_column=6)

        detail_cols = [("Analyst", 18), ("Brand", 22), ("ASIN", 14),
                       ("Net Price", 12), ("Recommended", 15), ("Status", 12)]
        write_header_row(ws, sep_row + 1, detail_cols, 1)

        view_cols_map = [
            (ac, 0), (bc, 1), (col_map.get("asin"), 2),
            (col_map.get("net_price"), 3), (rc, 4),
        ]
        dr = sep_row + 2
        for _, row in df_full.iterrows():
            rec_norm = row.get("_rec_norm", "Other")
            for src_col, ci in view_cols_map:
                if src_col and src_col in df_full.columns:
                    cell = ws.cell(row=dr, column=ci + 1, value=row[src_col])
                    cell.font = reg(9)
                    cell.fill = white_fill() if dr % 2 == 0 else light_fill()
                    cell.alignment = center()
                    cell.border = thin_border()

            status_cell = ws.cell(row=dr, column=6, value=rec_norm)
            if rec_norm == "Recommended":
                status_cell.fill = PatternFill("solid", fgColor="DCFCE7")
                status_cell.font = Font(name="Calibri", size=9, color=GREEN_HEX, bold=True)
            elif rec_norm == "Rejected":
                status_cell.fill = PatternFill("solid", fgColor="FEE2E2")
                status_cell.font = Font(name="Calibri", size=9, color=RED_HEX, bold=True)
            else:
                status_cell.fill = light_fill()
                status_cell.font = reg(9)
            status_cell.alignment = center()
            status_cell.border = thin_border()
            ws.row_dimensions[dr].height = 18
            dr += 1

    ws.freeze_panes = "A4"


# ── Sheet 3: Brand Intelligence ───────────────────────────────────────────────
def build_brand_sheet(wb, data):
    ws = wb.create_sheet("🏷️ Brand Intelligence")
    ws.sheet_view.showGridLines = False

    brand_df = data["brand_df"]
    bc_col   = data["bc_col"]
    df_full  = data["df_full"]
    col_map  = data["col_map"]

    ws.merge_cells("A1:H1")
    ws["A1"] = "BRAND PERFORMANCE INTELLIGENCE"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE_HEX)
    ws["A1"].fill = navy_fill()
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 8

    if not brand_df.empty and bc_col:
        name_col = bc_col if bc_col in brand_df.columns else brand_df.columns[0]
        total_approved = brand_df["Approved SKUs"].sum()

        brand_cols = [("Rank", 8), ("Brand", 28), ("Approved SKUs", 16),
                      ("Share of Approvals", 18), ("Priority", 12)]
        write_header_row(ws, 3, brand_cols, 1)

        for ri, (_, row) in enumerate(brand_df.iterrows()):
            r = 4 + ri
            count = int(row["Approved SKUs"])
            share = count / total_approved if total_approved else 0
            priority = "🔥 TOP" if ri == 0 else ("★ HIGH" if ri < 3 else "MID")
            row_fill = PatternFill("solid", fgColor="FFF7ED") if ri == 0 else (
                white_fill() if ri % 2 == 0 else light_fill())

            for ci, val in enumerate([ri + 1, row[name_col], count, f"{share*100:.1f}%", priority], 1):
                cell = ws.cell(row=r, column=ci, value=val)
                cell.font = bold_dark(10) if ri == 0 else reg(10)
                cell.fill = row_fill
                cell.alignment = center()
                cell.border = thin_border()
            ws.row_dimensions[r].height = 20

        # Add a bar chart
        chart_data_start = 4
        chart_data_end   = 4 + len(brand_df) - 1
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Approved SKUs by Brand"
        chart.y_axis.title = "Brand"
        chart.x_axis.title = "Approved SKUs"
        chart.style = 10
        chart.width  = 20
        chart.height = 14

        data_ref   = Reference(ws, min_col=3, min_row=chart_data_start, max_row=chart_data_end)
        cats_ref   = Reference(ws, min_col=2, min_row=chart_data_start, max_row=chart_data_end)
        chart.add_data(data_ref)
        chart.set_categories(cats_ref)
        chart.series[0].graphicalProperties.solidFill = ORANGE_HEX
        chart.series[0].graphicalProperties.line.solidFill = ORANGE_HEX

        ws.add_chart(chart, "F3")

    ws.freeze_panes = "A4"


# ── Sheet 4: Raw Data ─────────────────────────────────────────────────────────
def build_raw_data_sheet(wb, data):
    ws = wb.create_sheet("📄 Raw Data")
    ws.sheet_view.showGridLines = False

    df_full  = data["df_full"]
    view_cols = data["view_cols"]

    ws.merge_cells("A1:H1")
    ws["A1"] = "FULL AUDIT DATA — FILTERED VIEW"
    ws["A1"].font = Font(name="Calibri", bold=True, size=13, color=WHITE_HEX)
    ws["A1"].fill = navy_fill()
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 8

    display_df = df_full[view_cols].copy() if view_cols else df_full.copy()

    for ci, col in enumerate(display_df.columns, 1):
        cell = ws.cell(row=3, column=ci, value=str(col))
        cell.font = bold_white(10)
        cell.fill = header_fill()
        cell.alignment = center()
        cell.border = thin_border()
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.row_dimensions[3].height = 22

    for ri, (_, row) in enumerate(display_df.iterrows()):
        r = 4 + ri
        rec_norm = df_full.iloc[ri].get("_rec_norm", "Other") if "_rec_norm" in df_full.columns else "Other"
        row_fill = (PatternFill("solid", fgColor="F0FDF4") if rec_norm == "Recommended"
                    else (PatternFill("solid", fgColor="FFF1F2") if rec_norm == "Rejected"
                          else white_fill() if ri % 2 == 0 else light_fill()))
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=ci, value=val)
            cell.font = reg(9)
            cell.fill = row_fill
            cell.alignment = left()
            cell.border = thin_border()
        ws.row_dimensions[r].height = 17

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = ws.dimensions


# ── Main entry ─────────────────────────────────────────────────────────────────
def generate_excel(data: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    build_dashboard(wb, data)
    build_analyst_sheet(wb, data)
    build_brand_sheet(wb, data)
    build_raw_data_sheet(wb, data)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
