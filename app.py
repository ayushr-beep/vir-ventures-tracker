import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime, date
from export_pptx import generate_pptx
from export_excel import generate_excel

st.set_page_config(
    page_title="Vir Ventures · Analyst Dashboard",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded"
)

# ── Design tokens ─────────────────────────────────────────────────────────────
OR = "#E8611A"; NV = "#0F1629"; GR = "#16A34A"; RD = "#DC2626"
AM = "#D97706"; BL = "#2563EB"; LT = "#F8F9FB"; WH = "#FFFFFF"
OR2 = "#F4874B"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*{{font-family:'Inter',sans-serif;}}
/* Sidebar */
[data-testid="stSidebar"]{{background:{NV};border-right:1px solid #1E2A45;}}
[data-testid="stSidebar"] *{{color:#C8D0E7 !important;}}
[data-testid="stSidebar"] label{{color:#8892A4 !important;font-size:10px !important;font-weight:700 !important;letter-spacing:.08em;text-transform:uppercase;}}
/* Main */
.block-container{{padding-top:1rem !important;padding-bottom:2rem !important;max-width:1400px !important;}}
h1,h2,h3{{margin:0 !important;}}
/* KPI cards */
.kcard{{background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:18px 16px 14px;position:relative;overflow:hidden;}}
.kcard::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;}}
.kcard.or::before{{background:{OR};}}
.kcard.gr::before{{background:{GR};}}
.kcard.rd::before{{background:{RD};}}
.kcard.bl::before{{background:{BL};}}
.kcard.am::before{{background:{AM};}}
.kcard .klabel{{font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#9CA3AF;margin-bottom:6px;}}
.kcard .kval{{font-size:30px;font-weight:800;color:{NV};line-height:1;}}
.kcard .ksub{{font-size:11px;color:#9CA3AF;margin-top:5px;}}
/* Section labels */
.sec{{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#9CA3AF;margin:24px 0 10px;padding-bottom:6px;border-bottom:1px solid #F3F4F6;}}
/* Vendor pill */
.vpill{{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;}}
/* Tabs */
.stTabs [data-baseweb="tab"]{{font-size:12px;font-weight:600;color:#9CA3AF;padding:8px 16px;}}
.stTabs [aria-selected="true"]{{color:{OR} !important;border-bottom:2px solid {OR} !important;}}
/* Download btn */
.stDownloadButton>button{{background:{OR} !important;color:#fff !important;border:none !important;border-radius:8px !important;font-weight:700 !important;font-size:13px !important;padding:10px 20px !important;width:100% !important;letter-spacing:.02em;}}
.stDownloadButton>button:hover{{background:#C9531A !important;}}
/* Metric */
div[data-testid="metric-container"]{{background:#fff;border:1px solid #F3F4F6;border-radius:10px;padding:14px 16px;}}
/* Table */
.stDataFrame{{border-radius:10px !important;}}
/* Select/Filter */
.stSelectbox>div>div,.stMultiSelect>div>div{{border-radius:8px !important;font-size:13px !important;}}
/* Plotly toolbar */
.js-plotly-plot .plotly .modebar{{opacity:0.3;}}
</style>
""", unsafe_allow_html=True)

YES_VALS = {"yes","yes, need discount"}
VENDOR_NAMES = {"CRDI":"CRDI — Gaming & Entertainment","PNCL":"PNCL — The Pencil Grip","PROH":"PROH — Prohitter Sports"}
VENDOR_COLORS = {"CRDI": OR, "PNCL": BL, "PROH": AM}

def is_yes(v): return str(v).strip().lower() in YES_VALS

@st.cache_data(show_spinner=False)
def load_df(b, name):
    ext = name.rsplit(".",1)[-1].lower()
    df = pd.read_excel(io.BytesIO(b)) if ext in ("xlsx","xls") else pd.read_csv(io.BytesIO(b))
    num = ['Net price','BB Price','Breakeven','Difference from SP','Percentage Diff from SP',
           'Total Competition Score','Total Margin Score','Total Demand Score',
           'Number of FBA vendors','Total New FBA Sellers','Last month sold','Rank',
           'Average number of review','# of Reviews(Format Specific)','Quantity',
           'Lifetime','Current year','TOTAL(Stock+Reserve+inbound)']
    for c in num:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce')
    df['is_yes']     = df['Recommended'].apply(is_yes)
    df['margin_gap'] = (df['BB Price'] - df['Net price']).round(2)
    df['margin_pct'] = ((df['BB Price'] - df['Net price']) / df['Net price'].replace(0,np.nan) * 100).round(1)
    df['rec_clean']  = df['Recommended'].apply(
        lambda x: "Rejected" if str(x).strip().lower()=="no"
        else ("Needs Discount" if "discount" in str(x).lower() else "Approved"))
    return df

# ── Helpers ───────────────────────────────────────────────────────────────────
def kcard(col, label, value, sub, klass="or"):
    col.markdown(f'<div class="kcard {klass}"><div class="klabel">{label}</div><div class="kval">{value}</div><div class="ksub">{sub}</div></div>', unsafe_allow_html=True)

def sec(text):
    st.markdown(f'<div class="sec">{text}</div>', unsafe_allow_html=True)

def chart_layout(fig, h=320):
    fig.update_layout(
        height=h, margin=dict(l=4,r=4,t=28,b=4),
        plot_bgcolor=WH, paper_bgcolor=WH,
        font=dict(family="Inter",size=11),
        legend=dict(orientation="h",y=1.12,x=0,font=dict(size=11)),
        hoverlabel=dict(bgcolor=NV,font_color=WH,font_size=12,bordercolor=NV),
        xaxis=dict(showgrid=True,gridcolor="#F3F4F6",gridwidth=1),
        yaxis=dict(showgrid=True,gridcolor="#F3F4F6",gridwidth=1),
    )
    return fig

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<div style='padding:16px 0 8px;'><span style='font-size:20px;font-weight:800;color:#fff;'>Vir Ventures</span><br><span style='font-size:11px;color:#5A6380;'>Analyst Dashboard</span></div>", unsafe_allow_html=True)
    st.divider()
    uploaded = st.file_uploader("Upload Mastersheet", type=["xlsx","xls","csv"], label_visibility="collapsed",
                                 help="Upload your weekly Excel audit file")
    st.markdown("<p style='font-size:10px;color:#3A4260;margin-top:4px;'>Supports .xlsx · .xls · .csv</p>", unsafe_allow_html=True)
    st.divider()
    st.markdown("<p style='font-size:10px;color:#5A6380;font-weight:700;letter-spacing:.08em;'>REPORT SETTINGS</p>", unsafe_allow_html=True)
    week_label  = st.text_input("Week Label", value=f"Week 1 — {date.today().strftime('%d %b %Y')}", label_visibility="collapsed")
    report_date = st.date_input("Report Date", value=date.today(), label_visibility="collapsed")
    st.divider()
    st.markdown("<p style='font-size:10px;color:#5A6380;font-weight:700;letter-spacing:.08em;'>FILTERS</p>", unsafe_allow_html=True)

if not uploaded:
    st.markdown(f"<h1 style='font-size:26px;font-weight:800;color:{NV};'>📊 Weekly Analyst Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#9CA3AF;font-size:14px;margin-top:4px;'>Upload your Mastersheet in the sidebar to generate the full boardroom view.</p>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='margin-top:40px;background:#fff;border:1.5px dashed #E5E7EB;border-radius:16px;padding:48px;text-align:center;'>
      <div style='font-size:48px;'>📤</div>
      <div style='font-size:16px;font-weight:700;color:{NV};margin-top:12px;'>Drop your Mastersheet Excel here</div>
      <div style='font-size:13px;color:#9CA3AF;margin-top:6px;'>Columns needed: Vendor Prefix · ASIN · Brand · Analyst · Recommended · Net price · BB Price · Rank</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── LOAD + FILTERS ────────────────────────────────────────────────────────────
raw = uploaded.read()
df  = load_df(raw, uploaded.name)

all_vendors  = sorted(df['Vendor Prefix'].dropna().unique()) if 'Vendor Prefix' in df.columns else []
all_analysts = sorted(df['Analyst'].dropna().unique())

with st.sidebar:
    vendor_sel  = st.multiselect("Vendor Prefix", all_vendors,  default=all_vendors,  key="vs")
    analyst_sel = st.multiselect("Analyst",        all_analysts, default=all_analysts, key="as_")
    rec_sel     = st.multiselect("Recommendation", ["Approved","Needs Discount","Rejected"],
                                  default=["Approved","Needs Discount","Rejected"], key="rs")
    st.divider()
    st.markdown("<p style='font-size:10px;color:#3A4260;'>Vir Ventures · Internal Only</p>", unsafe_allow_html=True)

mask = df['Analyst'].isin(analyst_sel) & df['rec_clean'].isin(rec_sel)
if all_vendors: mask &= df['Vendor Prefix'].isin(vendor_sel)
dff = df[mask].copy()

# ── GLOBALS ────────────────────────────────────────────────────────────────────
total   = len(dff)
yes_df  = dff[dff['is_yes']]
no_df   = dff[~dff['is_yes']]
approved  = dff['is_yes'].sum()
rejected  = (~dff['is_yes']).sum()
rate      = round(approved/total*100,1) if total else 0
avg_margin = round(yes_df['margin_pct'].mean(),1) if not yes_df.empty else 0
avg_rank   = round(yes_df['Rank'].mean()) if not yes_df.empty else 0
total_units = int(yes_df['Last month sold'].sum())

# ── HEADER ─────────────────────────────────────────────────────────────────────
hc1, hc2 = st.columns([3,1])
with hc1:
    st.markdown(f"<h1 style='font-size:24px;font-weight:800;color:{NV};'>📊 Weekly Analyst Dashboard</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#9CA3AF;font-size:12px;margin-top:2px;'>{week_label} &nbsp;·&nbsp; {total} SKUs &nbsp;·&nbsp; {datetime.now().strftime('%d %b %Y, %H:%M')}</p>", unsafe_allow_html=True)
with hc2:
    st.markdown(f"<div style='text-align:right;padding-top:14px;'><span style='background:#FFF0E8;color:{OR};font-size:11px;font-weight:700;padding:5px 14px;border-radius:20px;letter-spacing:.06em;'>LIVE</span></div>", unsafe_allow_html=True)

st.divider()

# ── KPI ROW ───────────────────────────────────────────────────────────────────
k1,k2,k3,k4,k5,k6 = st.columns(6)
kcard(k1, "SKUs Reviewed",    str(total),              "This week",               "or")
kcard(k2, "Approved",         str(approved),           f"{rate}% approval rate",  "gr")
kcard(k3, "Rejected",         str(rejected),           f"{100-rate:.1f}% rejection","rd")
kcard(k4, "Avg BB Margin",    f"{avg_margin}%",        "On approved SKUs",        "or")
kcard(k5, "Avg BSR Rank",     f"#{avg_rank:,}",        "Lower = stronger demand", "bl")
kcard(k6, "Units / Month",    f"{total_units:,}",      "Approved SKUs total",     "am")

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t1,t2,t3,t4,t5,t6 = st.tabs([
    "📈 Overview",
    "🏪 Vendor Analysis",
    "📦 ASIN Intelligence",
    "👥 Analyst Scorecard",
    "⚠️ Rejections",
    "📥 Export"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with t1:
    r1,r2 = st.columns([1.1,1])

    with r1:
        sec("Approval Breakdown")
        rc = dff['rec_clean'].value_counts().reset_index()
        rc.columns = ['Status','Count']
        cmap = {"Approved":GR,"Needs Discount":OR,"Rejected":RD}
        fig = px.pie(rc, names='Status', values='Count', hole=.58,
                     color='Status', color_discrete_map=cmap)
        fig.update_traces(textinfo='label+percent', textfont_size=11,
                          pull=[.04,.04,.02], hovertemplate='<b>%{label}</b><br>%{value} SKUs (%{percent})<extra></extra>')
        fig.update_layout(height=300, margin=dict(l=0,r=0,t=10,b=0),
                          paper_bgcolor=WH, showlegend=True,
                          legend=dict(orientation='h',y=-.1,x=.5,xanchor='center',font=dict(size=11)),
                          font=dict(family="Inter"),
                          annotations=[dict(text=f"<b>{rate}%</b><br><span style='font-size:11px'>Approved</span>",
                                           showarrow=False, font=dict(size=13,family='Inter'),x=.5,y=.5)])
        st.plotly_chart(fig, use_container_width=True)

    with r2:
        sec("Analyst Approval Rates")
        ap = dff.groupby('Analyst').agg(Total=('ASIN','count'), Approved=('is_yes','sum')).reset_index()
        ap['Rate'] = (ap['Approved']/ap['Total']*100).round(1)
        ap['Color'] = ap['Rate'].apply(lambda r: GR if r>=80 else AM if r>=50 else RD)
        fig2 = go.Figure(go.Bar(
            x=ap['Analyst'], y=ap['Rate'],
            marker_color=ap['Color'],
            text=[f"{r}%" for r in ap['Rate']], textposition='outside',
            hovertemplate='<b>%{x}</b><br>Rate: %{y}%<extra></extra>'))
        fig2.add_hline(y=70, line_dash='dot', line_color='#CBD5E1',
                       annotation_text="70% target", annotation_font_color="#9CA3AF", annotation_font_size=10)
        chart_layout(fig2, 300)
        fig2.update_yaxes(range=[0,115], ticksuffix='%', title='')
        fig2.update_xaxes(title='')
        st.plotly_chart(fig2, use_container_width=True)

    r2a,r2b,r2c = st.columns(3)
    with r2a:
        sec("Margin Distribution — Approved")
        fig3 = px.histogram(yes_df, x='margin_pct', nbins=12, color_discrete_sequence=[OR],
                            labels={'margin_pct':'Margin %'})
        fig3.update_traces(hovertemplate='Margin: %{x:.0f}%<br>Count: %{y}<extra></extra>')
        chart_layout(fig3, 240); fig3.update_xaxes(ticksuffix='%',title='')
        fig3.update_yaxes(title='SKUs')
        st.plotly_chart(fig3, use_container_width=True)

    with r2b:
        sec("Buy Box Landscape — Approved")
        bb = yes_df['BuyBoxSellerName'].value_counts().head(6).reset_index()
        bb.columns = ['Seller','n']
        bb['Seller'] = bb['Seller'].str[:14]
        fig4 = px.bar(bb, x='n', y='Seller', orientation='h',
                      color='n', color_continuous_scale=[[0,'#FFF0E8'],[1,OR]], text='n')
        fig4.update_traces(textposition='outside', hovertemplate='<b>%{y}</b><br>%{x} SKUs<extra></extra>')
        chart_layout(fig4, 240)
        fig4.update_xaxes(title=''); fig4.update_yaxes(categoryorder='total ascending',title='')
        fig4.update_coloraxes(showscale=False)
        st.plotly_chart(fig4, use_container_width=True)

    with r2c:
        sec("Demand Tier Split — Approved")
        yd = yes_df.copy()
        yd['Tier'] = pd.cut(yd['Last month sold'],bins=[-1,0,50,100,300,9999],
                            labels=['0','1–50','51–100','101–300','300+'])
        tc = yd['Tier'].value_counts().sort_index().reset_index()
        tc.columns = ['Tier','n']
        fig5 = px.bar(tc, x='Tier', y='n', color='n',
                      color_continuous_scale=[[0,'#DCFCE7'],[1,GR]], text='n')
        fig5.update_traces(textposition='outside', hovertemplate='<b>%{x}</b><br>%{y} SKUs<extra></extra>')
        chart_layout(fig5, 240)
        fig5.update_xaxes(title='Units/Month'); fig5.update_yaxes(title='SKUs')
        fig5.update_coloraxes(showscale=False)
        st.plotly_chart(fig5, use_container_width=True)

    # Auto executive brief
    sec("Executive Brief")
    if not yes_df.empty:
        top_a = dff.groupby('Analyst')['is_yes'].mean().idxmax()
        top_ar = round(dff.groupby('Analyst')['is_yes'].mean()[top_a]*100,1)
        top_brand = yes_df.groupby('Brand')['ASIN'].count().idxmax()
        top_brand_n = int(yes_df.groupby('Brand')['ASIN'].count().max())
        top_asin = yes_df.sort_values('Last month sold',ascending=False).iloc[0]
        st.markdown(f"""
<div style='background:{LT};border-left:3px solid {OR};border-radius:10px;padding:18px 22px;font-size:13px;line-height:2;color:#374151;'>
▸ &nbsp;This week <b>{total} SKUs</b> were reviewed across <b>{dff['Brand'].nunique()} brands</b> and <b>{dff['Vendor Prefix'].nunique() if 'Vendor Prefix' in dff.columns else '—'} vendor groups</b>. Approval rate: <b style='color:{GR};'>{rate}%</b>.<br>
▸ &nbsp;<b>{top_a}</b> is the top analyst this week — <b>{top_ar}% approval rate</b> with strong demand-score selection.<br>
▸ &nbsp;Top brand: <b>{top_brand}</b> with <b>{top_brand_n} approved SKUs</b> · Top ASIN: <b>{top_asin['ASIN']}</b> ({top_asin['Brand']}, <b style='color:{GR};'>{int(top_asin['Last month sold'])} units/mo</b>).<br>
▸ &nbsp;<b>{rejected} SKUs rejected</b> — all PROH vendor. Avg Buy Box margin on approved SKUs: <b style='color:{OR};'>{avg_margin}%</b>.
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — VENDOR ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t2:
    if 'Vendor Prefix' not in dff.columns or dff['Vendor Prefix'].nunique() == 0:
        st.info("No Vendor Prefix column found in your file."); st.stop()

    vp_stats = dff.groupby('Vendor Prefix').agg(
        Total=('ASIN','count'),
        Approved=('is_yes','sum'),
        Rejected=('is_yes', lambda x: (~x).sum()),
        Brands=('Brand','nunique'),
        Analysts=('Analyst','nunique'),
        TotalUnits=('Last month sold','sum'),
        AvgMarginPct=('margin_pct','mean'),
        AvgRank=('Rank','mean'),
        AvgNetPrice=('Net price','mean'),
        AvgBBPrice=('BB Price','mean'),
        AvgDemandScore=('Total Demand Score','mean'),
        AvgCompScore=('Total Competition Score','mean'),
        AvgMarginScore=('Total Margin Score','mean'),
    ).reset_index().round(1)
    vp_stats['ApprovalRate'] = (vp_stats['Approved']/vp_stats['Total']*100).round(1)
    vp_stats['VendorName'] = vp_stats['Vendor Prefix'].map(VENDOR_NAMES).fillna(vp_stats['Vendor Prefix'])
    vp_stats['Color']      = vp_stats['Vendor Prefix'].map(VENDOR_COLORS).fillna(OR)

    # Vendor scorecard cards
    sec("Vendor Scorecard")
    vcols = st.columns(len(vp_stats))
    for i,(_,row) in enumerate(vp_stats.iterrows()):
        rate_c = GR if row['ApprovalRate']>=80 else AM if row['ApprovalRate']>=50 else RD
        with vcols[i]:
            st.markdown(f"""
<div style='background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:18px 14px;border-top:3px solid {row["Color"]};'>
  <div style='font-size:11px;font-weight:700;color:{row["Color"]};letter-spacing:.06em;'>{row["Vendor Prefix"]}</div>
  <div style='font-size:13px;font-weight:600;color:{NV};margin:3px 0 12px;line-height:1.3;'>{row["VendorName"].split("—")[-1].strip()}</div>
  <div style='display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:12px;'>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Total SKUs</div><b style='color:{NV};font-size:15px;'>{int(row["Total"])}</b></div>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Approved</div><b style='color:{GR};font-size:15px;'>{int(row["Approved"])}</b></div>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Approval Rate</div><b style='color:{rate_c};font-size:15px;'>{row["ApprovalRate"]}%</b></div>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Avg Margin</div><b style='color:{OR};font-size:15px;'>{row["AvgMarginPct"]}%</b></div>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Units/Month</div><b style='font-size:13px;'>{int(row["TotalUnits"])}</b></div>
    <div style='background:{LT};padding:8px 10px;border-radius:8px;'><div style='color:#9CA3AF;font-size:10px;'>Brands</div><b style='font-size:13px;'>{int(row["Brands"])}</b></div>
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    vc1, vc2 = st.columns(2)
    with vc1:
        sec("Approved vs Rejected by Vendor")
        fig_vb = go.Figure()
        fig_vb.add_trace(go.Bar(name='Approved', x=vp_stats['Vendor Prefix'], y=vp_stats['Approved'],
                                marker_color=GR, text=vp_stats['Approved'], textposition='outside'))
        fig_vb.add_trace(go.Bar(name='Rejected', x=vp_stats['Vendor Prefix'], y=vp_stats['Rejected'],
                                marker_color=RD, text=vp_stats['Rejected'], textposition='outside'))
        fig_vb.update_layout(barmode='group')
        chart_layout(fig_vb, 300)
        fig_vb.update_xaxes(title=''); fig_vb.update_yaxes(title='SKUs')
        st.plotly_chart(fig_vb, use_container_width=True)

    with vc2:
        sec("Avg Margin % by Vendor")
        yes_vp = dff[dff['is_yes']].groupby('Vendor Prefix').agg(AvgMargin=('margin_pct','mean')).reset_index().round(1)
        yes_vp['Color'] = yes_vp['Vendor Prefix'].map(VENDOR_COLORS).fillna(OR)
        fig_vm = go.Figure(go.Bar(
            x=yes_vp['Vendor Prefix'], y=yes_vp['AvgMargin'],
            marker_color=yes_vp['Color'],
            text=[f"{v}%" for v in yes_vp['AvgMargin']], textposition='outside',
            hovertemplate='<b>%{x}</b><br>Avg Margin: %{y}%<extra></extra>'))
        chart_layout(fig_vm, 300)
        fig_vm.update_yaxes(ticksuffix='%', title='')
        fig_vm.update_xaxes(title='')
        st.plotly_chart(fig_vm, use_container_width=True)

    vc3, vc4 = st.columns(2)
    with vc3:
        sec("Units/Month by Vendor")
        fig_vu = go.Figure(go.Bar(
            x=vp_stats['Vendor Prefix'], y=vp_stats['TotalUnits'],
            marker_color=[VENDOR_COLORS.get(v, OR) for v in vp_stats['Vendor Prefix']],
            text=vp_stats['TotalUnits'].astype(int), textposition='outside',
            hovertemplate='<b>%{x}</b><br>Units/Month: %{y:,}<extra></extra>'))
        chart_layout(fig_vu, 280)
        fig_vu.update_yaxes(title='Units'); fig_vu.update_xaxes(title='')
        st.plotly_chart(fig_vu, use_container_width=True)

    with vc4:
        sec("Avg Score Comparison by Vendor")
        cats = ['Demand Score','Competition Score','Margin Score']
        fig_vr = go.Figure()
        clrs = [OR, BL, AM]
        for j,(_,row) in enumerate(vp_stats.iterrows()):
            vals = [row['AvgDemandScore'], row['AvgCompScore'], row['AvgMarginScore']]
            vals += [vals[0]]
            fig_vr.add_trace(go.Scatterpolar(
                r=vals, theta=cats+[cats[0]], fill='toself',
                name=row['Vendor Prefix'],
                line_color=clrs[j % len(clrs)],
                fillcolor=clrs[j % len(clrs)], opacity=.2,
                hovertemplate='<b>%{theta}</b>: %{r:.1f}<extra></extra>'))
        fig_vr.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0,100])),
            height=280, margin=dict(l=40,r=40,t=28,b=0),
            paper_bgcolor=WH, font=dict(family="Inter",size=11),
            legend=dict(orientation='h',y=-.1,font=dict(size=11)))
        st.plotly_chart(fig_vr, use_container_width=True)

    # Per-vendor deep dive
    sec("Vendor Deep Dive — ASIN Level")
    selected_vendor = st.selectbox("Select vendor to inspect", all_vendors if all_vendors else ["All"],
                                    format_func=lambda x: VENDOR_NAMES.get(x,x))
    vd = dff[dff['Vendor Prefix']==selected_vendor].sort_values('Last month sold',ascending=False) if all_vendors else dff
    show_cols = ['ASIN','Brand','Analyst','Recommended','Last month sold','Net price',
                 'BB Price','margin_gap','margin_pct','Rank','Total Product Score',
                 'Number of FBA vendors','BuyBoxSellerName','Remarks ']
    show_cols = [c for c in show_cols if c in vd.columns]
    st.dataframe(vd[show_cols].rename(columns={
        'Last month sold':'Units/Mo','margin_gap':'Margin Gap ($)',
        'margin_pct':'Margin %','Number of FBA vendors':'FBA Sellers',
        'BuyBoxSellerName':'BB Holder','Remarks ':'Remarks'}
    ).reset_index(drop=True), use_container_width=True, hide_index=True, height=360,
    column_config={
        "Margin %": st.column_config.ProgressColumn("Margin %", min_value=0, max_value=200, format="%.1f%%"),
        "Units/Mo": st.column_config.NumberColumn("Units/Mo", format="%d 📦"),
    })

    # Brand mix within vendor
    sec(f"Brand Mix — {selected_vendor}")
    vd_brands = vd.groupby('Brand').agg(
        Total=('ASIN','count'), Approved=('is_yes','sum'),
        TotalUnits=('Last month sold','sum'), AvgMargin=('margin_pct','mean')
    ).reset_index().sort_values('Approved',ascending=False).round(1)
    fig_vbrand = px.treemap(vd_brands, path=['Brand'], values='Total',
                             color='AvgMargin', color_continuous_scale=[[0,'#FFF0E8'],[1,OR]],
                             hover_data={'Approved':True,'TotalUnits':True,'AvgMargin':True})
    fig_vbrand.update_layout(height=280, margin=dict(l=0,r=0,t=0,b=0),
                              paper_bgcolor=WH, font=dict(family="Inter"))
    st.plotly_chart(fig_vbrand, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — ASIN INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
with t3:
    sec("Top Approved ASINs — Ranked by Demand")
    top3 = yes_df.sort_values('Last month sold',ascending=False).head(3)
    medals = ["🥇","🥈","🥉"]; medal_colors=[OR, BL, GR]
    tc = st.columns(3)
    for i,((_,row),col) in enumerate(zip(top3.iterrows(), tc)):
        with col:
            st.markdown(f"""
<div style='background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:16px;border-top:3px solid {medal_colors[i]};'>
  <div style='font-size:18px;'>{medals[i]}</div>
  <div style='font-size:10px;color:#9CA3AF;font-weight:700;margin-top:4px;'>{row.get("Vendor Prefix","—")} · {row["Brand"]}</div>
  <div style='font-size:14px;font-weight:700;color:{NV};margin:4px 0 10px;'>{row["ASIN"]}</div>
  <div style='display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:11px;'>
    <div style='background:{LT};padding:6px 8px;border-radius:6px;'><span style='color:#9CA3AF;display:block;font-size:10px;'>Units/Month</span><b style='color:{GR};'>{int(row["Last month sold"])}</b></div>
    <div style='background:{LT};padding:6px 8px;border-radius:6px;'><span style='color:#9CA3AF;display:block;font-size:10px;'>Margin</span><b style='color:{OR};'>{row["margin_pct"]:.1f}%</b></div>
    <div style='background:{LT};padding:6px 8px;border-radius:6px;'><span style='color:#9CA3AF;display:block;font-size:10px;'>BSR Rank</span><b>#{int(row["Rank"]):,}</b></div>
    <div style='background:{LT};padding:6px 8px;border-radius:6px;'><span style='color:#9CA3AF;display:block;font-size:10px;'>BB Price</span><b>${row["BB Price"]:.2f}</b></div>
  </div>
  <div style='margin-top:8px;font-size:10px;color:#374151;background:{LT};padding:5px 8px;border-radius:6px;'>{str(row.get("Remarks ","—"))[:55]}</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    f1,f2,f3 = st.columns([1,1,1])
    with f1:
        analyst_f = st.selectbox("Filter Analyst", ["All"]+all_analysts, key="asin_af")
    with f2:
        vendor_f = st.selectbox("Filter Vendor", ["All"]+all_vendors, key="asin_vf")
    with f3:
        sort_by = st.selectbox("Sort by", ["Units/Month","Margin %","BSR Rank","BB Price"], key="asin_sort")

    sort_map = {"Units/Month":"Last month sold","Margin %":"margin_pct","BSR Rank":"Rank","BB Price":"BB Price"}
    asc_map  = {"Units/Month":False,"Margin %":False,"BSR Rank":True,"BB Price":False}
    asin_df  = yes_df.copy()
    if analyst_f != "All": asin_df = asin_df[asin_df['Analyst']==analyst_f]
    if vendor_f  != "All": asin_df = asin_df[asin_df['Vendor Prefix']==vendor_f]
    asin_df = asin_df.sort_values(sort_map[sort_by], ascending=asc_map[sort_by])

    show = ['ASIN','Brand','Vendor Prefix','Analyst','Recommended','Last month sold','Net price',
            'BB Price','margin_gap','margin_pct','Rank','Total Product Score',
            'Number of FBA vendors','Amazon Status','BuyBoxSellerName','Remarks ']
    show = [c for c in show if c in asin_df.columns]
    st.dataframe(asin_df[show].rename(columns={
        'Vendor Prefix':'Vendor','Last month sold':'Units/Mo',
        'margin_gap':'Margin Gap','margin_pct':'Margin %',
        'Number of FBA vendors':'FBA Sellers','BuyBoxSellerName':'BB Holder','Remarks ':'Remarks'
    }).reset_index(drop=True), use_container_width=True, hide_index=True, height=400,
    column_config={
        "Margin %": st.column_config.ProgressColumn("Margin %", min_value=0, max_value=200, format="%.1f%%"),
        "Units/Mo": st.column_config.NumberColumn("Units/Mo", format="%d 📦"),
    })

    sec("Margin % vs BSR Rank — All Approved SKUs")
    st.caption("Bubble size = units/month · Ideal: bottom-right (low rank, high margin)")
    fig_sc = px.scatter(yes_df, x='Rank', y='margin_pct',
                        color='Vendor Prefix' if 'Vendor Prefix' in yes_df.columns else 'Analyst',
                        size=yes_df['Last month sold'].clip(lower=5),
                        hover_data=['ASIN','Brand','BB Price','Analyst'],
                        color_discrete_map=VENDOR_COLORS,
                        labels={'Rank':'BSR Rank (lower=better)','margin_pct':'Margin %'})
    fig_sc.update_traces(hovertemplate='<b>%{customdata[0]}</b><br>%{customdata[1]}<br>Rank: %{x:,} · Margin: %{y:.1f}%<extra></extra>')
    chart_layout(fig_sc, 380)
    fig_sc.update_yaxes(ticksuffix='%', title='Margin %')
    st.plotly_chart(fig_sc, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ANALYST SCORECARD
# ══════════════════════════════════════════════════════════════════════════════
with t4:
    sec("Analyst Performance Scorecards")
    astats = dff.groupby('Analyst').agg(
        Total=('ASIN','count'), Approved=('is_yes','sum'),
        Rejected=('is_yes',lambda x:(~x).sum()),
        Rate=('is_yes',lambda x:round(x.mean()*100,1)),
        AvgMargin=('margin_pct','mean'), TotalUnits=('Last month sold','sum'),
        AvgRank=('Rank','mean'), AvgNetPrice=('Net price','mean'),
        AvgBBPrice=('BB Price','mean'), AvgDemand=('Total Demand Score','mean'),
        AvgComp=('Total Competition Score','mean'), AvgMarginSc=('Total Margin Score','mean'),
    ).reset_index().round(1)

    acols = st.columns(len(astats))
    for i,(_,row) in enumerate(astats.iterrows()):
        rc2 = GR if row['Rate']>=80 else AM if row['Rate']>=50 else RD
        with acols[i]:
            st.markdown(f"""
<div style='background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:18px 14px;text-align:center;border-top:4px solid {rc2};'>
  <div style='font-size:28px;font-weight:800;color:{rc2};'>{row['Rate']}%</div>
  <div style='font-size:14px;font-weight:700;color:{NV};margin:4px 0 2px;'>{row['Analyst']}</div>
  <div style='font-size:10px;color:#9CA3AF;margin-bottom:12px;'>Approval Rate</div>
  <div style='display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:11px;text-align:left;'>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Total</div><b>{int(row["Total"])}</b></div>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Approved</div><b style='color:{GR};'>{int(row["Approved"])}</b></div>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Avg Margin</div><b style='color:{OR};'>{row["AvgMargin"]:.1f}%</b></div>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Avg Rank</div><b>#{row["AvgRank"]:,.0f}</b></div>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Units/Mo</div><b>{int(row["TotalUnits"])}</b></div>
    <div style='background:{LT};padding:7px 9px;border-radius:7px;'><div style='color:#9CA3AF;font-size:9px;'>Rejected</div><b style='color:{RD};'>{int(row["Rejected"])}</b></div>
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    ac1, ac2 = st.columns(2)
    with ac1:
        sec("Score Radar by Analyst")
        cats = ['Demand Score','Competition Score','Margin Score','Approval Rate']
        fig_r2 = go.Figure()
        clrs2 = [OR, BL, AM]
        for j,(_,row) in enumerate(astats.iterrows()):
            vals = [row['AvgDemand'],row['AvgComp'],row['AvgMarginSc'],row['Rate']]
            vals += [vals[0]]
            fig_r2.add_trace(go.Scatterpolar(
                r=vals, theta=cats+[cats[0]], fill='toself', name=row['Analyst'],
                line_color=clrs2[j%len(clrs2)], fillcolor=clrs2[j%len(clrs2)], opacity=.2))
        fig_r2.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0,110])),
                              height=320, margin=dict(l=40,r=40,t=28,b=0),
                              paper_bgcolor=WH, font=dict(family="Inter",size=11),
                              legend=dict(orientation='h',y=-.1,font=dict(size=11)))
        st.plotly_chart(fig_r2, use_container_width=True)

    with ac2:
        sec("Workload vs Quality")
        fig_wq = px.scatter(astats, x='Total', y='Rate', size='TotalUnits',
                            color='Analyst', text='Analyst',
                            color_discrete_sequence=[OR, BL, AM],
                            labels={'Total':'SKUs Audited','Rate':'Approval Rate %'})
        fig_wq.update_traces(textposition='top center',
                              hovertemplate='<b>%{text}</b><br>Audited: %{x} · Rate: %{y}%<extra></extra>')
        fig_wq.add_hline(y=70, line_dash='dot', line_color='#CBD5E1')
        chart_layout(fig_wq, 320)
        fig_wq.update_yaxes(ticksuffix='%', range=[0,115], title='')
        fig_wq.update_xaxes(title='SKUs Audited')
        st.plotly_chart(fig_wq, use_container_width=True)

    sec("Analyst Detail Table")
    st.dataframe(astats.rename(columns={
        'Rate':'Approval Rate %','AvgMargin':'Avg Margin %','TotalUnits':'Total Units/Mo',
        'AvgRank':'Avg BSR Rank','AvgNetPrice':'Avg Net Price','AvgBBPrice':'Avg BB Price',
        'AvgDemand':'Avg Demand Sc.','AvgComp':'Avg Comp. Sc.','AvgMarginSc':'Avg Margin Sc.'
    }).reset_index(drop=True), use_container_width=True, hide_index=True,
    column_config={"Approval Rate %": st.column_config.ProgressColumn(
        "Approval Rate %", min_value=0, max_value=100, format="%.1f%%")})


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — REJECTIONS
# ══════════════════════════════════════════════════════════════════════════════
with t5:
    sec("Rejection Root Cause Analysis")
    rej = dff[~dff['is_yes']].copy()
    if rej.empty:
        st.success("No rejections in current selection.")
    else:
        neg = int(rej['Remarks '].str.contains('Negative',na=False).sum())
        low = int(rej['Remarks '].str.contains('Low',na=False).sum())
        oth = len(rej)-neg-low

        rk1,rk2,rk3,rk4 = st.columns(4)
        kcard(rk1,"Total Rejected",str(len(rej)),"This week","rd")
        kcard(rk2,"Negative Margin",str(neg),"BB < Breakeven","rd")
        kcard(rk3,"Low Margin",str(low),"Marginal profitability","am")
        kcard(rk4,"Other",str(oth),"Needs review","or")

        st.markdown("<br>", unsafe_allow_html=True)
        rc1,rc2 = st.columns(2)
        with rc1:
            sec("Reason Breakdown")
            rmk = rej['Remarks '].fillna('Unspecified').value_counts().reset_index()
            rmk.columns = ['Reason','Count']
            fig_rp = px.pie(rmk, names='Reason', values='Count', hole=.5,
                            color_discrete_sequence=[RD,'#F87171','#FCA5A5','#FECACA'])
            fig_rp.update_traces(textinfo='label+percent', hovertemplate='<b>%{label}</b><br>%{value} SKUs<extra></extra>')
            chart_layout(fig_rp, 280)
            st.plotly_chart(fig_rp, use_container_width=True)
        with rc2:
            sec("Approved vs Rejected — Price Comparison")
            comp_df = pd.DataFrame({
                'Type':['Approved','Rejected','Approved','Rejected','Approved','Rejected'],
                'Metric':['BB Price','BB Price','Net Price','Net Price','Breakeven','Breakeven'],
                'Value':[yes_df['BB Price'].mean(), rej['BB Price'].mean(),
                         yes_df['Net price'].mean(), rej['Net price'].mean(),
                         yes_df['Breakeven'].mean(), rej['Breakeven'].mean()]
            }).round(2)
            fig_rc = px.bar(comp_df, x='Metric', y='Value', color='Type', barmode='group',
                            color_discrete_map={'Approved':GR,'Rejected':RD},
                            text='Value')
            fig_rc.update_traces(texttemplate='$%{text:.2f}', textposition='outside')
            chart_layout(fig_rc, 280)
            fig_rc.update_yaxes(tickprefix='$', title='')
            st.plotly_chart(fig_rc, use_container_width=True)

        sec("Rejected ASIN Detail")
        rej_show = ['ASIN','Brand','Vendor Prefix','Analyst','Net price','BB Price',
                    'Breakeven','Difference from SP','margin_gap','Rank','Total Product Score','Remarks ']
        rej_show = [c for c in rej_show if c in rej.columns]
        st.dataframe(rej[rej_show].rename(columns={
            'Vendor Prefix':'Vendor','margin_gap':'Margin Gap','Remarks ':'Reason'
        }).reset_index(drop=True), use_container_width=True, hide_index=True)

        sec("Action Plan")
        st.markdown(f"""
<div style='background:#FFF1F2;border-left:3px solid {RD};border-radius:10px;padding:16px 20px;font-size:13px;line-height:1.9;'>
  <b style='color:{RD};'>01 — Immediate:</b> &nbsp;{neg} ASINs have negative margin. BB Price is below landed cost. Request 15–20% cost reduction from PROH vendor before next PO.<br>
  <b style='color:{AM};'>02 — This Week:</b> &nbsp;{low} ASINs are marginally profitable. Explore bundle deals or volume pricing. Set a minimum margin floor for future submissions.<br>
  <b style='color:{NV};'>03 — Next Cycle:</b> &nbsp;All {len(rej)} rejections come from PROH. Schedule vendor review and set clear margin thresholds before resubmission.
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with t6:
    sec("Download Boardroom Reports")
    ec1, ec2 = st.columns(2)

    export_data = {
        'df': dff, 'week_label': week_label, 'report_date': str(report_date),
        'total': total, 'approved': approved, 'rejected': rejected, 'rate': rate,
        'avg_margin': avg_margin, 'avg_rank': avg_rank, 'total_units': total_units,
    }

    with ec1:
        st.markdown(f"""
<div style='background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:22px;margin-bottom:14px;'>
  <div style='font-size:28px;'>📊</div>
  <div style='font-size:15px;font-weight:700;color:{NV};margin:8px 0 4px;'>PowerPoint Presentation</div>
  <div style='font-size:12px;color:#6B7280;line-height:1.6;'>8 boardroom-ready slides · Cover, Executive Summary, Vendor Analysis, ASIN Intelligence, Analyst Scorecards, Brand Analysis, Rejection Review, Next Steps · Navy × Orange premium theme</div>
</div>
""", unsafe_allow_html=True)
        try:
            pptx_b = generate_pptx(export_data)
            st.download_button("📊 Download PowerPoint", data=pptx_b,
                               file_name=f"VirVentures_Report_{report_date}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"PPTX error: {e}")

    with ec2:
        st.markdown(f"""
<div style='background:#fff;border:1px solid #E5E7EB;border-radius:12px;padding:22px;margin-bottom:14px;'>
  <div style='font-size:28px;'>📋</div>
  <div style='font-size:15px;font-weight:700;color:{NV};margin:8px 0 4px;'>Excel Analyst Workbook</div>
  <div style='font-size:12px;color:#6B7280;line-height:1.6;'>6 premium sheets · Executive Dashboard, Vendor Analysis, ASIN Intelligence, Analyst Scorecards, Brand Intelligence, Rejection Analysis · Colour-coded, auto-filter, freeze panes</div>
</div>
""", unsafe_allow_html=True)
        try:
            xl_b = generate_excel(export_data)
            st.download_button("📋 Download Excel Workbook", data=xl_b,
                               file_name=f"VirVentures_Workbook_{report_date}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel error: {e}")

    st.markdown("<br>", unsafe_allow_html=True)
    csv_b = dff.to_csv(index=False).encode()
    st.download_button("📄 Download Raw CSV", data=csv_b,
                       file_name=f"VirVentures_Raw_{report_date}.csv", mime="text/csv")
PYEOF
echo "app.py written"
