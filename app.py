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
    page_title="Vir Ventures · Boardroom Analyst Dashboard",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
[data-testid="stSidebar"]{background:#0F1629;border-right:1px solid #1E2A45;}
[data-testid="stSidebar"] *{color:#C8D0E7 !important;}
[data-testid="stSidebar"] h2,h3{color:#fff !important;}
[data-testid="stSidebar"] label{color:#8892A4 !important;font-size:11px !important;font-weight:600 !important;letter-spacing:.06em;text-transform:uppercase;}
[data-testid="stSidebar"] .stFileUploader{border:1.5px dashed #2D3A5C !important;border-radius:10px !important;background:#131D33 !important;}
.block-container{padding-top:1.2rem !important;max-width:1400px !important;}
h1{font-size:24px !important;font-weight:800 !important;color:#0F1629 !important;}
.stDownloadButton>button{background:#E8611A !important;color:#fff !important;border:none !important;border-radius:10px !important;font-weight:700 !important;padding:11px 22px !important;width:100%;}
.stDownloadButton>button:hover{background:#C9531A !important;}
.stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600;color:#6B7280;}
.stTabs [aria-selected="true"]{color:#E8611A !important;border-bottom:2px solid #E8611A !important;}
div[data-testid="metric-container"]{background:#fff;border:1px solid #E5E7EB;border-radius:14px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.06);}
</style>
""", unsafe_allow_html=True)

ORANGE = "#E8611A"; NAVY = "#0F1629"; GREEN = "#16A34A"; RED = "#DC2626"
AMBER = "#D97706"; BLUE = "#2563EB"; LIGHT = "#F8F9FB"
YES_VALS = {"yes","yes, need discount"}

def is_yes(v): return str(v).strip().lower() in YES_VALS

@st.cache_data(show_spinner=False)
def load_df(b, name):
    ext = name.rsplit(".",1)[-1].lower()
    df = pd.read_excel(io.BytesIO(b)) if ext in ("xlsx","xls") else pd.read_csv(io.BytesIO(b))
    num = ['Net price','BB Price','Breakeven','Difference from SP','Percentage Diff from SP',
           'Total Competition Score','Total Margin Score','Total Demand Score',
           'Number of FBA vendors','Total New FBA Sellers','Last month sold','Rank',
           'Average number of review','# of Reviews(Format Specific)','Quantity',
           'Lifetime','Current year','TOTAL(Stock+Reserve+inbound)','Total Product Rating']
    for c in num:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce')
    df['is_yes']      = df['Recommended'].apply(is_yes)
    df['margin_gap']  = (df['BB Price'] - df['Net price']).round(2)
    df['margin_pct']  = ((df['BB Price'] - df['Net price']) / df['Net price'].replace(0,np.nan) * 100).round(1)
    df['rec_clean']   = df['Recommended'].apply(lambda x: "Recommended" if is_yes(x) else ("Needs Discount" if "discount" in str(x).lower() else "Rejected"))
    df['rec_clean']   = df['Recommended'].apply(
        lambda x: "Rejected" if str(x).strip().lower()=="no"
        else ("Needs Discount" if "discount" in str(x).lower() else "Approved"))
    return df

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📦 Vir Ventures")
    st.markdown("<p style='font-size:11px;color:#5A6380;margin-top:-10px;'>Boardroom Analyst Dashboard</p>", unsafe_allow_html=True)
    st.markdown("---")
    uploaded = st.file_uploader("Upload Mastersheet", type=["xlsx","xls","csv"])
    st.markdown("---")
    st.markdown("### Report Settings")
    week_label  = st.text_input("Week Label", value=f"Week 1 — {date.today().strftime('%d %b %Y')}")
    report_date = st.date_input("Report Date", value=date.today())
    st.markdown("---")
    st.markdown("### Filters")
    analyst_filter = st.multiselect("Analyst", options=[], key="af")
    prefix_filter  = st.multiselect("Vendor Prefix", options=[], key="pf")
    st.markdown("---")
    st.markdown("<p style='font-size:10px;color:#3A4260;'>Vir Ventures · Internal Use Only</p>", unsafe_allow_html=True)

# ── NO FILE ───────────────────────────────────────────────────────────────────
if not uploaded:
    st.markdown("# 📊 Boardroom Analyst Dashboard")
    st.markdown("<p style='color:#6B7280;margin-top:-10px;'>Upload your weekly Mastersheet to generate the full executive view.</p>", unsafe_allow_html=True)
    st.info("👈 Upload your Mastersheet Excel file in the sidebar to get started.")
    st.stop()

# ── LOAD ──────────────────────────────────────────────────────────────────────
raw = uploaded.read()
df  = load_df(raw, uploaded.name)

# Populate sidebar filters
all_analysts = sorted(df['Analyst'].dropna().unique())
all_prefixes = sorted(df['Vendor Prefix'].dropna().unique()) if 'Vendor Prefix' in df.columns else []

with st.sidebar:
    analyst_filter = st.multiselect("Analyst", options=all_analysts, default=all_analysts, key="af2")
    prefix_filter  = st.multiselect("Vendor Prefix", options=all_prefixes, default=all_prefixes, key="pf2")

mask = df['Analyst'].isin(analyst_filter)
if all_prefixes: mask &= df['Vendor Prefix'].isin(prefix_filter)
df = df[mask].copy()

# ── METRICS ───────────────────────────────────────────────────────────────────
total      = len(df)
approved   = df['is_yes'].sum()
rejected   = (~df['is_yes']).sum()
rate       = round(approved/total*100,1) if total else 0
avg_margin = df.loc[df['is_yes'],'margin_pct'].mean()
avg_rank   = df.loc[df['is_yes'],'Rank'].mean()
total_demand = df.loc[df['is_yes'],'Last month sold'].sum()

# ── HEADER ────────────────────────────────────────────────────────────────────
c1,c2 = st.columns([3,1])
with c1:
    st.markdown(f"# 📊 Boardroom Analyst Dashboard")
    st.markdown(f"<p style='color:#6B7280;font-size:13px;margin-top:-14px;'>{week_label} &nbsp;·&nbsp; {total} SKUs reviewed &nbsp;·&nbsp; Generated {datetime.now().strftime('%d %b %Y %H:%M')}</p>", unsafe_allow_html=True)
with c2:
    st.markdown("<div style='text-align:right;padding-top:10px;'><span style='background:#FFF0E8;color:#E8611A;font-size:11px;font-weight:700;padding:5px 14px;border-radius:20px;letter-spacing:.06em;'>WEEK 1 LIVE</span></div>", unsafe_allow_html=True)

st.markdown("---")

# ── KPI ROW ───────────────────────────────────────────────────────────────────
k1,k2,k3,k4,k5,k6 = st.columns(6)
def kpi(col, label, val, delta=None, help=None):
    col.metric(label, val, delta=delta, help=help)

kpi(k1, "📦 SKUs Reviewed",    f"{total}",           help="Total ASINs in this week's mastersheet")
kpi(k2, "✅ Approved",         f"{approved}",         f"{rate}% approval rate")
kpi(k3, "❌ Rejected",         f"{rejected}",         f"{100-rate:.1f}% rejection rate")
kpi(k4, "📈 Avg Buy Box Margin", f"{avg_margin:.1f}%" if not np.isnan(avg_margin) else "—", help="Avg margin % on approved SKUs")
kpi(k5, "🏆 Avg BSR (Approved)", f"{avg_rank:,.0f}" if not np.isnan(avg_rank) else "—",    help="Lower = better seller rank")
kpi(k6, "💰 Total Demand (units)", f"{int(total_demand):,}", help="Sum of Last Month Sold across approved SKUs")

st.markdown("<br>", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
tabs = st.tabs(["🎯 Executive Overview","📦 ASIN Intelligence","👥 Analyst Scorecard","🏷️ Brand & Vendor","⚠️ Rejected Deep-Dive","📥 Export Reports"])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — EXECUTIVE OVERVIEW
# ─────────────────────────────────────────────────────────────────────────────
with tabs[0]:
    r1c1, r1c2 = st.columns([1,1])

    with r1c1:
        st.markdown("#### Approval Breakdown")
        rec_counts = df['rec_clean'].value_counts().reset_index()
        rec_counts.columns = ['Status','Count']
        color_map = {"Approved":"#16A34A","Needs Discount":"#E8611A","Rejected":"#DC2626"}
        fig = px.pie(rec_counts, names='Status', values='Count', hole=.55,
                     color='Status', color_discrete_map=color_map)
        fig.update_traces(textinfo='label+percent', textfont_size=12, pull=[.04,.04,.02])
        fig.update_layout(height=300, margin=dict(l=0,r=0,t=0,b=0), paper_bgcolor='white',
                          showlegend=True, legend=dict(orientation='h',y=-.08,x=.5,xanchor='center'),
                          annotations=[dict(text=f"<b>{rate}%</b><br>Approved", showarrow=False,
                                           font=dict(size=14,family='Inter'),x=.5,y=.5)])
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        st.markdown("#### Approval Rate by Analyst")
        analyst_perf = df.groupby('Analyst').agg(
            Total=('ASIN','count'),
            Approved=('is_yes','sum'),
        ).reset_index()
        analyst_perf['Rate'] = (analyst_perf['Approved']/analyst_perf['Total']*100).round(1)
        colors = [GREEN if r>=80 else AMBER if r>=50 else RED for r in analyst_perf['Rate']]
        fig2 = go.Figure(go.Bar(
            x=analyst_perf['Analyst'], y=analyst_perf['Rate'],
            marker_color=colors, text=[f"{r}%" for r in analyst_perf['Rate']],
            textposition='outside'))
        fig2.add_hline(y=70, line_dash='dash', line_color='#9CA3AF',
                       annotation_text="70% target", annotation_position="top right")
        fig2.update_layout(height=300, margin=dict(l=0,r=0,t=10,b=0),
                           plot_bgcolor='white', paper_bgcolor='white',
                           yaxis=dict(range=[0,115],ticksuffix='%',showgrid=True,gridcolor='#F3F4F6'),
                           font=dict(family='Inter'), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        st.markdown("#### Margin Gap Distribution (Approved)")
        yes_df = df[df['is_yes']]
        fig3 = px.histogram(yes_df, x='margin_pct', nbins=15, color_discrete_sequence=[ORANGE],
                            labels={'margin_pct':'Margin % (BB vs Net Cost)'})
        fig3.update_layout(height=240, margin=dict(l=0,r=0,t=10,b=0),
                           plot_bgcolor='white', paper_bgcolor='white',
                           yaxis=dict(showgrid=True,gridcolor='#F3F4F6'),
                           font=dict(family='Inter'))
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        st.markdown("#### Buy Box Seller Landscape")
        bb = df[df['is_yes']]['BuyBoxSellerName'].value_counts().head(6).reset_index()
        bb.columns = ['Seller','Count']
        fig4 = px.bar(bb, x='Count', y='Seller', orientation='h',
                      color='Count', color_continuous_scale=[[0,'#FFF0E8'],[1,ORANGE]],
                      text='Count')
        fig4.update_traces(textposition='outside')
        fig4.update_layout(height=240, margin=dict(l=0,r=20,t=10,b=0),
                           plot_bgcolor='white', paper_bgcolor='white',
                           coloraxis_showscale=False,
                           yaxis=dict(categoryorder='total ascending'),
                           font=dict(family='Inter'))
        st.plotly_chart(fig4, use_container_width=True)

    with r2c3:
        st.markdown("#### Demand Tier Split (Approved)")
        yes_df2 = df[df['is_yes']].copy()
        yes_df2['Demand Tier'] = pd.cut(yes_df2['Last month sold'],
            bins=[-1,0,50,100,300,99999],
            labels=['0 units','1–50','51–100','101–300','300+'])
        tier = yes_df2['Demand Tier'].value_counts().sort_index().reset_index()
        tier.columns = ['Tier','Count']
        fig5 = px.bar(tier, x='Tier', y='Count', color='Count',
                      color_continuous_scale=[[0,'#DCFCE7'],[1,GREEN]], text='Count')
        fig5.update_layout(height=240, margin=dict(l=0,r=0,t=10,b=0),
                           plot_bgcolor='white', paper_bgcolor='white',
                           coloraxis_showscale=False, font=dict(family='Inter'))
        st.plotly_chart(fig5, use_container_width=True)

    # Executive Brief
    st.markdown("---")
    st.markdown("#### 📝 Executive Brief — Auto-Generated")
    top_analyst = df.groupby('Analyst')['is_yes'].mean().idxmax()
    top_analyst_rate = round(df.groupby('Analyst')['is_yes'].mean()[top_analyst]*100,1)
    top_brand = df[df['is_yes']].groupby('Brand')['ASIN'].count().idxmax()
    top_brand_n = int(df[df['is_yes']].groupby('Brand')['ASIN'].count().max())
    top_asin = df[df['is_yes']].sort_values('Last month sold',ascending=False).iloc[0]
    avg_m = df[df['is_yes']]['margin_pct'].mean()
    brief = f"""
**Week in Review — {week_label}**

▸ This week, Vir Ventures reviewed **{total} SKUs** across **{df['Brand'].nunique()} brands** and **{df['Vendor Prefix'].nunique() if 'Vendor Prefix' in df.columns else 'multiple'} vendor prefixes**. The overall approval rate stands at **{rate}%**, with **{approved} SKUs** approved for catalogue and **{rejected} SKUs** rejected due to margin or demand issues.

▸ **{top_analyst}** delivered the strongest performance this week with a **{top_analyst_rate}% approval rate**, demonstrating rigorous SKU selection. All approved SKUs from this analyst show strong Buy Box positioning with low FBA competition.

▸ **{top_brand}** leads brand performance with **{top_brand_n} approved SKUs**, making it the priority vendor for this week's sourcing push. Average Buy Box margin across approved SKUs is **{avg_m:.1f}%**, well above breakeven.

▸ Top demand ASIN this week is **{top_asin['ASIN']}** ({top_asin['Brand']}) with **{int(top_asin['Last month sold'])} units sold last month** — recommend immediate PO actioning.

▸ **{rejected} ASINs** were rejected, all flagged as Negative Margin or Low Margin. These require vendor renegotiation before reconsideration.
"""
    st.markdown(f'<div style="background:#FAFBFF;border:1px solid #E5E7EB;border-radius:14px;padding:20px 24px;font-size:13px;line-height:1.9;color:#1F2937;">{brief}</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — ASIN INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────
with tabs[1]:
    st.markdown("#### 🏆 Top Approved ASINs — Ranked by Demand")

    yes_df = df[df['is_yes']].copy().sort_values('Last month sold', ascending=False)

    # Score card row for top 3
    top3 = yes_df.head(3)
    cols = st.columns(3)
    medal = ["🥇","🥈","🥉"]
    for i,((_,row),col) in enumerate(zip(top3.iterrows(), cols)):
        with col:
            st.markdown(f"""
<div style="background:#fff;border:1px solid #E5E7EB;border-radius:14px;padding:16px;border-top:3px solid {ORANGE};">
<div style="font-size:20px;margin-bottom:4px;">{medal[i]}</div>
<div style="font-size:11px;color:#6B7280;font-weight:600;text-transform:uppercase;letter-spacing:.05em;">{row['Brand']}</div>
<div style="font-size:15px;font-weight:700;color:#0F1629;margin:4px 0;">{row['ASIN']}</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:10px;font-size:12px;">
  <div><span style="color:#6B7280;">Last Month Sold</span><br><b style="color:{GREEN};">{int(row['Last month sold'])} units</b></div>
  <div><span style="color:#6B7280;">Margin</span><br><b style="color:{ORANGE};">{row['margin_pct']:.1f}%</b></div>
  <div><span style="color:#6B7280;">BSR Rank</span><br><b>#{int(row['Rank']):,}</b></div>
  <div><span style="color:#6B7280;">BB Price</span><br><b>${row['BB Price']:.2f}</b></div>
</div>
<div style="margin-top:8px;font-size:11px;background:#F0FDF4;color:#166534;padding:4px 8px;border-radius:6px;">{str(row.get('Remarks ','—'))[:60]}</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Full approved ASIN table
    st.markdown("#### Full Approved ASIN Table")
    disp_cols = ['ASIN','Brand','Analyst','Recommended','Last month sold','Net price','BB Price',
                 'margin_gap','margin_pct','Rank','Total Product Score','Number of FBA vendors',
                 'Amazon Status','BuyBoxSellerName','Remarks ']
    disp = yes_df[[c for c in disp_cols if c in yes_df.columns]].copy()
    disp.columns = [c.strip() for c in disp.columns]
    disp = disp.rename(columns={
        'margin_gap':'Margin Gap ($)','margin_pct':'Margin %',
        'Last month sold':'Units/Month','Number of FBA vendors':'FBA Sellers',
        'BuyBoxSellerName':'Buy Box Holder'
    })

    st.dataframe(disp.reset_index(drop=True), use_container_width=True, hide_index=True,
                 column_config={
                     "Margin %": st.column_config.ProgressColumn("Margin %", min_value=0, max_value=200, format="%.1f%%"),
                     "Units/Month": st.column_config.NumberColumn("Units/Month", format="%d 📦"),
                 }, height=400)

    st.markdown("<br>")
    # Margin vs Rank scatter
    st.markdown("#### Margin % vs BSR Rank — Approved SKUs")
    st.caption("Lower Rank = higher demand. Ideal ASINs: top-left (low rank, high margin)")
    fig_s = px.scatter(yes_df, x='Rank', y='margin_pct', color='Analyst',
                       size='Last month sold', hover_data=['ASIN','Brand','BB Price'],
                       color_discrete_sequence=[ORANGE, BLUE, GREEN],
                       labels={'Rank':'BSR Rank (lower=better)','margin_pct':'Margin %'})
    fig_s.update_layout(height=380, plot_bgcolor='white', paper_bgcolor='white',
                        font=dict(family='Inter'),
                        yaxis=dict(showgrid=True,gridcolor='#F3F4F6',ticksuffix='%'),
                        xaxis=dict(showgrid=True,gridcolor='#F3F4F6'))
    st.plotly_chart(fig_s, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — ANALYST SCORECARD
# ─────────────────────────────────────────────────────────────────────────────
with tabs[2]:
    st.markdown("#### 👥 Analyst Performance Scorecard")

    analyst_stats = df.groupby('Analyst').agg(
        Total=('ASIN','count'),
        Approved=('is_yes','sum'),
        Rejected=('is_yes', lambda x: (~x).sum()),
        Approval_Rate=('is_yes', lambda x: round(x.mean()*100,1)),
        Avg_Rank=('Rank','mean'),
        Avg_Net_Price=('Net price','mean'),
        Avg_BB_Price=('BB Price','mean'),
        Avg_Margin_Pct=('margin_pct','mean'),
        Avg_Demand_Score=('Total Demand Score','mean'),
        Avg_Competition_Score=('Total Competition Score','mean'),
        Avg_Margin_Score=('Total Margin Score','mean'),
        Total_Units_Month=('Last month sold','sum'),
    ).reset_index().round(1)

    # Scorecards
    cols = st.columns(len(analyst_stats))
    for i,((_,row),col) in enumerate(zip(analyst_stats.iterrows(), cols)):
        rate_c = GREEN if row['Approval_Rate']>=80 else AMBER if row['Approval_Rate']>=50 else RED
        with col:
            st.markdown(f"""
<div style="background:#fff;border:1px solid #E5E7EB;border-radius:16px;padding:20px 16px;text-align:center;border-top:4px solid {rate_c};">
<div style="font-size:26px;font-weight:800;color:{rate_c};">{row['Approval_Rate']}%</div>
<div style="font-size:15px;font-weight:700;color:#0F1629;margin:4px 0;">{row['Analyst']}</div>
<div style="font-size:11px;color:#6B7280;margin-bottom:12px;">Approval Rate</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:11px;text-align:left;">
  <div style="background:#F8F9FB;padding:6px 8px;border-radius:6px;"><span style="color:#9CA3AF;">Total</span><br><b>{int(row['Total'])}</b></div>
  <div style="background:#F8F9FB;padding:6px 8px;border-radius:6px;"><span style="color:#9CA3AF;">Approved</span><br><b style="color:{GREEN};">{int(row['Approved'])}</b></div>
  <div style="background:#F8F9FB;padding:6px 8px;border-radius:6px;"><span style="color:#9CA3AF;">Avg Margin</span><br><b style="color:{ORANGE};">{row['Avg_Margin_Pct']:.1f}%</b></div>
  <div style="background:#F8F9FB;padding:6px 8px;border-radius:6px;"><span style="color:#9CA3AF;">Avg Rank</span><br><b>#{row['Avg_Rank']:,.0f}</b></div>
</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Score Comparison Radar")
        categories = ['Demand Score','Competition Score','Margin Score','Approval Rate']
        fig_r = go.Figure()
        colors_r = [ORANGE, BLUE, GREEN]
        for i, (_, row) in enumerate(analyst_stats.iterrows()):
            vals = [row['Avg_Demand_Score'], row['Avg_Competition_Score'],
                    row['Avg_Margin_Score'], row['Approval_Rate']]
            vals += [vals[0]]
            cats = categories + [categories[0]]
            fig_r.add_trace(go.Scatterpolar(r=vals, theta=cats, fill='toself',
                                            name=row['Analyst'], line_color=colors_r[i],
                                            fillcolor=colors_r[i], opacity=.25))
        fig_r.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0,110])),
                            height=340, paper_bgcolor='white',
                            font=dict(family='Inter'), legend=dict(orientation='h',y=-.1))
        st.plotly_chart(fig_r, use_container_width=True)

    with c2:
        st.markdown("#### Workload vs Quality Matrix")
        fig_m = px.scatter(analyst_stats, x='Total', y='Approval_Rate',
                           size='Total_Units_Month', color='Analyst',
                           text='Analyst',
                           color_discrete_sequence=[ORANGE,BLUE,GREEN],
                           labels={'Total':'SKUs Audited','Approval_Rate':'Approval Rate %'})
        fig_m.update_traces(textposition='top center')
        fig_m.add_hline(y=70, line_dash='dash', line_color='#9CA3AF')
        fig_m.update_layout(height=340, plot_bgcolor='white', paper_bgcolor='white',
                            font=dict(family='Inter'),
                            yaxis=dict(ticksuffix='%',range=[0,115],showgrid=True,gridcolor='#F3F4F6'),
                            xaxis=dict(showgrid=True,gridcolor='#F3F4F6'))
        st.plotly_chart(fig_m, use_container_width=True)

    st.markdown("#### Detailed Analyst Table")
    st.dataframe(analyst_stats.rename(columns={
        'Approval_Rate':'Approval Rate %','Avg_Rank':'Avg BSR Rank',
        'Avg_Net_Price':'Avg Net Price ($)','Avg_BB_Price':'Avg BB Price ($)',
        'Avg_Margin_Pct':'Avg Margin %','Avg_Demand_Score':'Avg Demand Score',
        'Avg_Competition_Score':'Avg Competition Score','Avg_Margin_Score':'Avg Margin Score',
        'Total_Units_Month':'Total Units/Month'
    }), use_container_width=True, hide_index=True,
    column_config={
        "Approval Rate %": st.column_config.ProgressColumn("Approval Rate %", min_value=0, max_value=100, format="%.1f%%"),
    })


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — BRAND & VENDOR
# ─────────────────────────────────────────────────────────────────────────────
with tabs[3]:
    st.markdown("#### 🏷️ Brand Intelligence")

    brand_stats = df.groupby('Brand').agg(
        Total_SKUs=('ASIN','count'),
        Approved=('is_yes','sum'),
        Avg_Margin_Pct=('margin_pct','mean'),
        Total_Units_Month=('Last month sold','sum'),
        Avg_Rank=('Rank','mean'),
        Avg_BB_Price=('BB Price','mean'),
        Avg_Net_Price=('Net price','mean'),
    ).reset_index().round(1)
    brand_stats['Approval_Rate'] = (brand_stats['Approved']/brand_stats['Total_SKUs']*100).round(1)
    brand_stats = brand_stats.sort_values('Approved', ascending=False)

    c1, c2 = st.columns([1.4,1])
    with c1:
        st.markdown("#### Brands by Approved SKUs")
        top_brands = brand_stats.head(12)
        fig_b = px.bar(top_brands, x='Approved', y='Brand', orientation='h',
                       text='Approved', color='Avg_Margin_Pct',
                       color_continuous_scale=[[0,'#FFF0E8'],[.5,ORANGE],[1,'#C9531A']],
                       labels={'Avg_Margin_Pct':'Avg Margin %'})
        fig_b.update_traces(textposition='outside')
        fig_b.update_layout(height=400, margin=dict(l=0,r=40,t=0,b=0),
                            plot_bgcolor='white', paper_bgcolor='white',
                            yaxis=dict(categoryorder='total ascending'),
                            coloraxis_colorbar=dict(title="Margin %",thickness=12,len=.7),
                            font=dict(family='Inter'))
        st.plotly_chart(fig_b, use_container_width=True)

    with c2:
        st.markdown("#### Demand vs Margin by Brand")
        fig_bm = px.scatter(brand_stats[brand_stats['Approved']>0],
                            x='Avg_Margin_Pct', y='Total_Units_Month',
                            size='Approved', color='Approval_Rate', text='Brand',
                            color_continuous_scale=[[0,RED],[.5,AMBER],[1,GREEN]],
                            labels={'Avg_Margin_Pct':'Avg Margin %','Total_Units_Month':'Total Units/Month'})
        fig_bm.update_traces(textposition='top center', textfont_size=9)
        fig_bm.update_layout(height=400, plot_bgcolor='white', paper_bgcolor='white',
                             font=dict(family='Inter'),
                             yaxis=dict(showgrid=True,gridcolor='#F3F4F6'),
                             xaxis=dict(showgrid=True,gridcolor='#F3F4F6',ticksuffix='%'),
                             coloraxis_colorbar=dict(title="App. Rate %",thickness=12))
        st.plotly_chart(fig_bm, use_container_width=True)

    if 'Vendor Prefix' in df.columns:
        st.markdown("#### Vendor Prefix Summary")
        vp = df.groupby('Vendor Prefix').agg(
            SKUs=('ASIN','count'), Approved=('is_yes','sum'),
            Brands=('Brand','nunique'), Total_Units=('Last month sold','sum'),
            Avg_Price=('Net price','mean'), Avg_Margin=('margin_pct','mean'),
        ).reset_index().round(1)
        vp['Approval Rate'] = (vp['Approved']/vp['SKUs']*100).round(1)
        st.dataframe(vp, use_container_width=True, hide_index=True)

    st.markdown("#### Full Brand Table")
    st.dataframe(brand_stats.rename(columns={
        'Total_SKUs':'Total SKUs','Avg_Margin_Pct':'Avg Margin %',
        'Total_Units_Month':'Total Units/Month','Avg_Rank':'Avg BSR Rank',
        'Avg_BB_Price':'Avg BB Price','Avg_Net_Price':'Avg Net Price',
        'Approval_Rate':'Approval Rate %'
    }), use_container_width=True, hide_index=True,
    column_config={
        "Approval Rate %": st.column_config.ProgressColumn("Approval Rate %", min_value=0, max_value=100, format="%.1f%%"),
    })


# ─────────────────────────────────────────────────────────────────────────────
# TAB 5 — REJECTED DEEP DIVE
# ─────────────────────────────────────────────────────────────────────────────
with tabs[4]:
    st.markdown("#### ⚠️ Rejected ASINs — Root Cause Analysis")
    rej = df[~df['is_yes']].copy()

    r1, r2 = st.columns(2)
    with r1:
        st.markdown("**Rejection Reason Breakdown**")
        remarks = rej['Remarks '].fillna('Unspecified').value_counts().reset_index()
        remarks.columns = ['Reason','Count']
        fig_rej = px.pie(remarks, names='Reason', values='Count', hole=.45,
                         color_discrete_sequence=[RED,'#F87171','#FCA5A5','#FECACA'])
        fig_rej.update_layout(height=280, paper_bgcolor='white', margin=dict(l=0,r=0,t=0,b=0),
                              font=dict(family='Inter'))
        st.plotly_chart(fig_rej, use_container_width=True)

    with r2:
        st.markdown("**Margin Analysis — Rejected vs Approved**")
        comp = pd.DataFrame({
            'Category': ['Approved','Rejected'],
            'Avg BB Price':   [df[df['is_yes']]['BB Price'].mean(),   rej['BB Price'].mean()],
            'Avg Net Price':  [df[df['is_yes']]['Net price'].mean(),  rej['Net price'].mean()],
            'Avg Breakeven':  [df[df['is_yes']]['Breakeven'].mean(),  rej['Breakeven'].mean()],
        }).round(2)
        fig_c = go.Figure()
        for col, color in [('Avg BB Price',ORANGE),('Avg Net Price',BLUE),('Avg Breakeven',RED)]:
            fig_c.add_trace(go.Bar(name=col, x=comp['Category'], y=comp[col],
                                   marker_color=color, text=comp[col].round(2), textposition='outside'))
        fig_c.update_layout(barmode='group', height=280, plot_bgcolor='white', paper_bgcolor='white',
                            margin=dict(l=0,r=0,t=10,b=0), font=dict(family='Inter'),
                            yaxis=dict(tickprefix='$',showgrid=True,gridcolor='#F3F4F6'),
                            legend=dict(orientation='h',y=1.1))
        st.plotly_chart(fig_c, use_container_width=True)

    st.markdown("**Rejected ASIN Detail**")
    rej_disp = rej[['ASIN','Brand','Analyst','Net price','BB Price','Breakeven',
                     'Difference from SP','Rank','Total Product Score','Remarks ']].copy()
    rej_disp['Margin Gap'] = (rej_disp['BB Price'] - rej_disp['Net price']).round(2)
    st.dataframe(rej_disp.reset_index(drop=True), use_container_width=True, hide_index=True)

    st.markdown("**Management Recommendation for Rejected ASINs**")
    neg_margin = (rej['Remarks '].str.contains('Negative',na=False)).sum()
    low_margin = (rej['Remarks '].str.contains('Low',na=False)).sum()
    st.markdown(f"""
<div style="background:#FFF1F2;border:1px solid #FECACA;border-radius:12px;padding:16px 20px;font-size:13px;line-height:1.8;">
▸ <b>{neg_margin} ASINs</b> have negative margins — BB Price is below breakeven. Recommend requesting <b>vendor price reduction of 15-20%</b> before reactivation.<br>
▸ <b>{low_margin} ASINs</b> show marginal profitability. Consider bundling or higher volume purchasing to improve unit economics.<br>
▸ All {len(rej)} rejected ASINs belong to <b>{rej['Brand'].nunique()} brands</b>. Schedule vendor review for next sourcing cycle.
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 6 — EXPORT
# ─────────────────────────────────────────────────────────────────────────────
with tabs[5]:
    st.markdown("#### 📥 Export Boardroom Reports")
    st.markdown("Both reports are generated live from your current filtered data.")

    export_data = {
        'df': df, 'week_label': week_label, 'report_date': str(report_date),
        'total': total, 'approved': approved, 'rejected': rejected, 'rate': rate,
        'avg_margin': avg_margin, 'avg_rank': avg_rank, 'total_demand': total_demand,
    }

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div style="background:#fff;border:1px solid #E5E7EB;border-radius:14px;padding:24px;">
<div style="font-size:32px;margin-bottom:8px;">📊</div>
<div style="font-size:16px;font-weight:700;color:#0F1629;margin-bottom:4px;">PowerPoint Presentation</div>
<div style="font-size:13px;color:#6B7280;margin-bottom:16px;">8 boardroom-ready slides. Cover, Executive Summary, ASIN Intelligence, Analyst Scorecards, Brand Analysis, Rejection Review, Sourcing Recommendations, and Next Steps. Navy × Orange premium theme.</div>
</div>
""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        try:
            pptx_bytes = generate_pptx(export_data)
            st.download_button("📊 Download PowerPoint (.pptx)", data=pptx_bytes,
                               file_name=f"VirVentures_BoardroomReport_{report_date}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"PPTX error: {e}")

    with c2:
        st.markdown("""
<div style="background:#fff;border:1px solid #E5E7EB;border-radius:14px;padding:24px;">
<div style="font-size:32px;margin-bottom:8px;">📋</div>
<div style="font-size:16px;font-weight:700;color:#0F1629;margin-bottom:4px;">Excel Analyst Workbook</div>
<div style="font-size:13px;color:#6B7280;margin-bottom:16px;">5 premium formatted sheets. Executive Dashboard, ASIN Intelligence, Analyst Scorecards, Brand Intelligence, and Rejection Analysis. Colour-coded rows, embedded charts, auto-filter, and freeze panes.</div>
</div>
""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        try:
            xl_bytes = generate_excel(export_data)
            st.download_button("📋 Download Excel Workbook (.xlsx)", data=xl_bytes,
                               file_name=f"VirVentures_AnalystWorkbook_{report_date}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel error: {e}")

    st.markdown("<br>")
    csv = df.to_csv(index=False).encode()
    st.download_button("📄 Download Raw Filtered CSV", data=csv,
                       file_name=f"VirVentures_RawData_{report_date}.csv", mime="text/csv")
