import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import re
from datetime import datetime, date
from export_pptx import generate_pptx
from export_excel import generate_excel

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Vir Ventures · Weekly Analyst Tracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #1A1F36;
    border-right: 1px solid #2D3350;
}
[data-testid="stSidebar"] * { color: #C8D0E7 !important; }
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #FFFFFF !important; }
[data-testid="stSidebar"] label { color: #A0A8C0 !important; font-size: 12px !important; font-weight: 500 !important; letter-spacing: 0.04em; }
[data-testid="stSidebar"] .stDateInput input,
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stSelectbox > div > div {
    background: #252A45 !important;
    border: 1px solid #3D4468 !important;
    color: #E8ECF4 !important;
    border-radius: 8px !important;
}

/* ── KPI Cards ── */
.kpi-card {
    background: #FFFFFF;
    border: 1px solid #EAEDF3;
    border-radius: 16px;
    padding: 24px 20px 20px;
    position: relative;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    height: 130px;
}
.kpi-card.accent { border-top: 4px solid #E8611A; }
.kpi-card.success { border-top: 4px solid #16A34A; }
.kpi-card.danger  { border-top: 4px solid #DC2626; }
.kpi-card.info    { border-top: 4px solid #2563EB; }
.kpi-label { font-size: 11px; font-weight: 600; letter-spacing: 0.07em; text-transform: uppercase; color: #8892A4; margin-bottom: 8px; }
.kpi-value { font-size: 36px; font-weight: 700; color: #111827; line-height: 1; }
.kpi-sub   { font-size: 12px; color: #9CA3AF; margin-top: 6px; }
.kpi-badge { position: absolute; top: 20px; right: 18px; font-size: 18px; }

/* ── Section headers ── */
.section-header {
    display: flex; align-items: center; gap: 10px;
    margin: 32px 0 16px;
    padding-bottom: 10px;
    border-bottom: 2px solid #F3F4F6;
}
.section-header span { font-size: 16px; font-weight: 600; color: #111827; }
.section-pill { background: #FFF0E8; color: #E8611A; font-size: 11px; font-weight: 600; padding: 2px 10px; border-radius: 20px; letter-spacing: 0.05em; }

/* ── Tables ── */
.stDataFrame { border-radius: 12px !important; overflow: hidden !important; }

/* ── Upload area ── */
.upload-box {
    background: #FAFBFF;
    border: 2px dashed #D1D8F0;
    border-radius: 14px;
    padding: 32px;
    text-align: center;
    transition: all 0.2s;
}
.upload-title { font-size: 15px; font-weight: 600; color: #374151; margin-bottom: 6px; }
.upload-sub   { font-size: 13px; color: #6B7280; }

/* ── Download buttons ── */
.stDownloadButton > button {
    width: 100%;
    background: #E8611A !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    padding: 12px 20px !important;
    letter-spacing: 0.02em;
    transition: background 0.2s !important;
}
.stDownloadButton > button:hover { background: #C9531A !important; }

/* ── Metric delta ── */
[data-testid="stMetricDelta"] { font-size: 12px !important; }

/* ── General ── */
.block-container { padding-top: 1.5rem !important; max-width: 1300px !important; }
h1 { font-size: 26px !important; font-weight: 700 !important; color: #111827 !important; }
h2 { font-size: 18px !important; font-weight: 600 !important; color: #1F2937 !important; }
h3 { font-size: 14px !important; font-weight: 600 !important; color: #374151 !important; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
ORANGE   = "#E8611A"
NAVY     = "#1A1F36"
YES_VALS = {"yes", "yes, low qty", "yes,low qty", "yes low qty"}
NO_VALS  = {"no"}

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalise_rec(val):
    if pd.isna(val): return "Other"
    v = str(val).strip().lower()
    if v in YES_VALS: return "Recommended"
    if v in NO_VALS:  return "Rejected"
    return "Other"

def rate_color(r):
    if r >= 70: return "🟢"
    if r >= 40: return "🟡"
    return "🔴"

@st.cache_data(show_spinner=False)
def load_data(uploaded_bytes, file_name):
    ext = file_name.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(io.BytesIO(uploaded_bytes))
    elif ext == "csv":
        df = pd.read_csv(io.BytesIO(uploaded_bytes))
    else:
        return None, "Unsupported file type"
    return df, None

def detect_columns(df):
    cols = {c.lower().strip(): c for c in df.columns}
    def find(candidates):
        for c in candidates:
            if c in cols: return cols[c]
        return None
    return {
        "recommended": find(["recommended", "an", "recom", "rec", "recommendation"]),
        "analyst":     find(["analysist", "analyst", "ar", "analysis", "analyst name"]),
        "brand":       find(["brand", "q", "brand name", "vendor"]),
        "asin":        find(["output asin", "asin", "p", "output_asin"]),
        "prefix":      find(["prefix"]),
        "net_price":   find(["sum of net price", "net price", "net_price", "n", "price"]),
        "traffic":     find(["traffic", "sessions", "page views", "pageviews"]),
        "sales":       find(["sales", "revenue", "ordered revenue", "gross sales"]),
    }

def compute_kpis(df, col_map):
    rc = col_map["recommended"]
    if not rc: return {}
    df = df.copy()
    df["_rec_norm"] = df[rc].apply(normalise_rec)
    total      = len(df)
    recommended = (df["_rec_norm"] == "Recommended").sum()
    rejected    = (df["_rec_norm"] == "Rejected").sum()
    other       = total - recommended - rejected
    rate        = round(recommended / total * 100, 1) if total else 0
    return dict(total=total, recommended=recommended, rejected=rejected,
                other=other, rate=rate, df=df)

def analyst_stats(df_norm, col_map):
    ac = col_map["analyst"]
    if not ac: return pd.DataFrame()
    g = df_norm.groupby(ac).apply(lambda x: pd.Series({
        "Total Audited":   len(x),
        "Recommended":     (x["_rec_norm"] == "Recommended").sum(),
        "Rejected":        (x["_rec_norm"] == "Rejected").sum(),
        "Approval Rate %": round((x["_rec_norm"] == "Recommended").sum() / len(x) * 100, 1) if len(x) else 0,
    })).reset_index()
    ac_name = ac if ac in ac else ac
    ac_col = ac
    if ac_col not in analyst_stats.__dict__:
        pass
    ac_col_name = ac
    ac_col_name_in_result = df_norm[ac].name
    return ac_col_name_in_result, ac_col, g

def brand_stats(df_norm, col_map, top_n=12):
    bc = col_map["brand"]
    if not bc: return pd.DataFrame()
    sub = df_norm[df_norm["_rec_norm"] == "Recommended"]
    s = sub.groupby(bc).size().reset_index(name="Approved SKUs")
    return s.sort_values("Approved SKUs", ascending=False).head(top_n)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Vir Ventures")
    st.markdown("**Weekly Analyst Tracker**")
    st.markdown("---")

    st.markdown("### 📁 Data Upload")
    uploaded = st.file_uploader(
        "Upload weekly audit file",
        type=["xlsx", "xls", "csv"],
        help="Upload your weekly product audit Excel or CSV file",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("### ⚙️ Report Settings")
    week_label = st.text_input("Week Label", value=f"Week of {date.today().strftime('%d %b %Y')}")
    report_date = st.date_input("Report Date", value=date.today())
    category_filter = st.text_input("Category Filter (optional)", placeholder="e.g. BLFN, SPQT")

    st.markdown("---")
    st.markdown("### 📥 Exports")
    export_placeholder = st.empty()

    st.markdown("---")
    st.markdown(
        "<div style='font-size:11px;color:#5A6380;'>Vir Ventures Internal Tool<br>Analyst Performance Dashboard</div>",
        unsafe_allow_html=True,
    )

# ── Main Header ───────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"# 📦 Weekly Analyst Performance Tracker")
    st.markdown(
        f"<p style='color:#6B7280;font-size:14px;margin-top:-12px;'>"
        f"{week_label} &nbsp;·&nbsp; Generated {datetime.now().strftime('%d %b %Y, %H:%M')}</p>",
        unsafe_allow_html=True,
    )
with col_h2:
    st.markdown(
        "<div style='text-align:right;padding-top:8px;'>"
        "<span style='background:#FFF0E8;color:#E8611A;font-size:12px;font-weight:700;"
        "padding:6px 14px;border-radius:20px;letter-spacing:0.05em;'>LIVE DASHBOARD</span></div>",
        unsafe_allow_html=True,
    )

st.markdown("---")

# ── No file state ─────────────────────────────────────────────────────────────
if not uploaded:
    st.markdown("""
    <div class="upload-box" style="margin-top:40px;">
        <div style="font-size:48px;margin-bottom:16px;">📤</div>
        <div class="upload-title">Upload your weekly audit file to get started</div>
        <div class="upload-sub">Supported formats: Excel (.xlsx, .xls) · CSV<br>
        Key columns needed: <strong>Recommended</strong> · <strong>Analyst</strong> · <strong>Brand</strong></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 Expected Data Structure")
    sample_cols = ["Prefix", "Brand", "Output ASIN", "Net Price", "Recommended", "Analysist", "Sales", "Traffic"]
    sample_data = [
        ["BLFN", "BANDAI", "B09Y9KZT3J", 7.70, "Yes", "Babul", 1240, 3500],
        ["BLFN", "Bandai Namco", "B0CBD7F8Q9", 12.00, "Yes", "Babul", 980, 2100],
        ["BLFN", "Dragon Ball Super", "B079YQH8MM", 15.40, "Yes, low qty", "Babul", 640, 1800],
        ["BLFN", "Tamagotchi", "B01N00XfI", 20.00, "No", "Chanchal", 210, 900],
        ["SPQT", "Spoontiques", "B07514Q6", 9.00, "Yes", "Chanchal", 1890, 5200],
        ["ZBRA", "Zebra Pen", "B00V7TK9Y", 7.00, "Yes", "Sunanda", 760, 2400],
    ]
    sample_df = pd.DataFrame(sample_data, columns=sample_cols)
    st.dataframe(sample_df, use_container_width=True, hide_index=True)
    st.caption("The tool auto-detects column names — 'Analysist', 'Analyst', 'AR' all work.")
    st.stop()

# ── Load & process ─────────────────────────────────────────────────────────────
raw_bytes = uploaded.read()
df_raw, err = load_data(raw_bytes, uploaded.name)
if err:
    st.error(f"Error loading file: {err}")
    st.stop()

col_map = detect_columns(df_raw)

# Category filter
if category_filter.strip():
    prefixes = [p.strip().upper() for p in category_filter.split(",")]
    pc = col_map["prefix"]
    if pc and pc in df_raw.columns:
        df_raw = df_raw[df_raw[pc].astype(str).str.upper().isin(prefixes)]

if not col_map["recommended"]:
    st.error("❌ Could not detect a 'Recommended' column. Please ensure your file has a column named 'Recommended', 'AN', or 'Rec'.")
    st.expander("Columns found in your file").write(list(df_raw.columns))
    st.stop()

# Compute
kpis = compute_kpis(df_raw, col_map)
df = kpis.pop("df")

ac_col = col_map["analyst"]
if ac_col:
    a_grouped = df.groupby(ac_col).apply(lambda x: pd.Series({
        "Total Audited":   len(x),
        "Recommended":     int((x["_rec_norm"] == "Recommended").sum()),
        "Rejected":        int((x["_rec_norm"] == "Rejected").sum()),
        "Approval Rate %": round((x["_rec_norm"] == "Recommended").sum() / len(x) * 100, 1) if len(x) else 0,
    })).reset_index()
    a_grouped = a_grouped.rename(columns={ac_col: "Analyst"})
    a_grouped = a_grouped.sort_values("Total Audited", ascending=False)
else:
    a_grouped = pd.DataFrame()

bc_col = col_map["brand"]
brand_df = brand_stats(df, col_map)

# ── KPI Cards ─────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
cards = [
    (c1, "CATALOGUE SIZE", f"{kpis['total']:,}", "Total SKUs reviewed", "accent", "📦"),
    (c2, "RECOMMENDED",    f"{kpis['recommended']:,}", "Yes + Yes, low qty", "success", "✅"),
    (c3, "REJECTED",       f"{kpis['rejected']:,}", "Flagged No", "danger", "❌"),
    (c4, "APPROVAL RATE",  f"{kpis['rate']}%", f"of {kpis['total']:,} items", "info", "📈"),
]
for col, label, value, sub, klass, icon in cards:
    with col:
        st.markdown(
            f'<div class="kpi-card {klass}">'
            f'<div class="kpi-badge">{icon}</div>'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'<div class="kpi-sub">{sub}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# ── Charts Row 1 ──────────────────────────────────────────────────────────────
st.markdown('<div class="section-header"><span>📊 Macro Overview</span><span class="section-pill">EXECUTIVE VIEW</span></div>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.4, 1])

with col_left:
    st.markdown("#### Top Brands by Approved SKUs")
    if not brand_df.empty:
        bc_label = bc_col if bc_col else "Brand"
        fig_bar = px.bar(
            brand_df, x="Approved SKUs", y=bc_col if bc_col in brand_df.columns else brand_df.columns[0],
            orientation="h", text="Approved SKUs",
            color="Approved SKUs",
            color_continuous_scale=[[0, "#FFF0E8"], [0.4, "#F4874B"], [1, "#E8611A"]],
        )
        fig_bar.update_traces(textposition="outside", textfont_size=11)
        fig_bar.update_layout(
            margin=dict(l=0, r=20, t=10, b=10), height=360,
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            yaxis=dict(categoryorder="total ascending", tickfont=dict(size=11), title=""),
            xaxis=dict(showgrid=True, gridcolor="#F3F4F6", title="Approved SKUs"),
            font=dict(family="Inter"),
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("No brand column detected.")

with col_right:
    st.markdown("#### Recommendation Breakdown")
    rec_counts = df["_rec_norm"].value_counts().reset_index()
    rec_counts.columns = ["Status", "Count"]
    colors_map = {"Recommended": "#16A34A", "Rejected": "#DC2626", "Other": "#9CA3AF"}
    fig_pie = px.pie(
        rec_counts, names="Status", values="Count",
        color="Status", color_discrete_map=colors_map,
        hole=0.55,
    )
    fig_pie.update_traces(textinfo="label+percent", textfont_size=12, pull=[0.03, 0.03, 0])
    fig_pie.update_layout(
        margin=dict(l=10, r=10, t=10, b=10), height=360,
        paper_bgcolor="white", showlegend=True,
        legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center", font=dict(size=11)),
        font=dict(family="Inter"),
        annotations=[dict(text=f"<b>{kpis['rate']}%</b><br>Approved", showarrow=False,
                          font=dict(size=14, family="Inter"), x=0.5, y=0.5)],
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ── Analyst Performance ───────────────────────────────────────────────────────
st.markdown('<div class="section-header"><span>👥 Analyst Performance</span><span class="section-pill">MICRO TRACKING</span></div>', unsafe_allow_html=True)

if not a_grouped.empty:
    col_tbl, col_chart = st.columns([1.1, 1])

    with col_tbl:
        st.markdown("#### Individual Performance Table")
        disp = a_grouped.copy()
        disp["Rate Signal"] = disp["Approval Rate %"].apply(
            lambda r: "🟢 High" if r >= 70 else ("🟡 Mid" if r >= 40 else "🔴 Low")
        )
        st.dataframe(
            disp[["Analyst", "Total Audited", "Recommended", "Rejected", "Approval Rate %", "Rate Signal"]],
            use_container_width=True, hide_index=True,
            column_config={
                "Approval Rate %": st.column_config.ProgressColumn(
                    "Approval Rate %", min_value=0, max_value=100, format="%.1f%%"
                ),
                "Rate Signal": "Signal",
            },
        )

    with col_chart:
        st.markdown("#### Audited vs Recommended per Analyst")
        fig_ana = go.Figure()
        fig_ana.add_trace(go.Bar(
            name="Total Audited", x=a_grouped["Analyst"], y=a_grouped["Total Audited"],
            marker_color="#E2E8F0", text=a_grouped["Total Audited"], textposition="outside"
        ))
        fig_ana.add_trace(go.Bar(
            name="Recommended", x=a_grouped["Analyst"], y=a_grouped["Recommended"],
            marker_color=ORANGE, text=a_grouped["Recommended"], textposition="outside"
        ))
        fig_ana.update_layout(
            barmode="group", height=340, margin=dict(l=0, r=0, t=10, b=10),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", y=1.05, font=dict(size=11)),
            xaxis=dict(tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor="#F3F4F6"),
            font=dict(family="Inter"),
        )
        st.plotly_chart(fig_ana, use_container_width=True)

    # Approval rate trend by analyst (if week col exists)
    st.markdown("#### Analyst Approval Rate Comparison")
    fig_rate = go.Figure()
    for _, row in a_grouped.iterrows():
        fig_rate.add_trace(go.Bar(
            x=[row["Analyst"]], y=[row["Approval Rate %"]],
            name=row["Analyst"], text=[f"{row['Approval Rate %']:.1f}%"],
            textposition="outside",
            marker_color=ORANGE if row["Approval Rate %"] == a_grouped["Approval Rate %"].max() else "#CBD5E1",
        ))
    fig_rate.add_hline(y=70, line_dash="dash", line_color="#16A34A",
                       annotation_text="70% target", annotation_position="right")
    fig_rate.update_layout(
        height=280, margin=dict(l=0, r=60, t=10, b=10),
        plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
        yaxis=dict(range=[0, 110], ticksuffix="%", showgrid=True, gridcolor="#F3F4F6"),
        xaxis=dict(tickfont=dict(size=12)),
        font=dict(family="Inter"),
    )
    st.plotly_chart(fig_rate, use_container_width=True)
else:
    st.info("No analyst column detected. Add an 'Analyst' or 'Analysist' column to see analyst performance.")

# ── Vendor Sales & Traffic ─────────────────────────────────────────────────────
sales_col   = col_map["sales"]
traffic_col = col_map["traffic"]

if sales_col or traffic_col:
    st.markdown('<div class="section-header"><span>💰 Vendor Sales & Traffic</span><span class="section-pill">MARKET INTELLIGENCE</span></div>', unsafe_allow_html=True)
    col_s1, col_s2 = st.columns(2)
    if sales_col and bc_col:
        with col_s1:
            st.markdown("#### Revenue by Brand (Top 10)")
            sales_brand = df.groupby(bc_col)[sales_col].sum().reset_index()
            sales_brand = sales_brand.sort_values(sales_col, ascending=False).head(10)
            fig_s = px.bar(sales_brand, x=bc_col, y=sales_col, text_auto=".2s",
                           color_discrete_sequence=[ORANGE])
            fig_s.update_layout(height=300, margin=dict(l=0, r=0, t=10, b=40),
                                 plot_bgcolor="white", paper_bgcolor="white",
                                 xaxis_tickangle=-30, font=dict(family="Inter"),
                                 yaxis=dict(showgrid=True, gridcolor="#F3F4F6"))
            st.plotly_chart(fig_s, use_container_width=True)
    if traffic_col and bc_col:
        with col_s2:
            st.markdown("#### Traffic (Sessions) by Brand (Top 10)")
            traffic_brand = df.groupby(bc_col)[traffic_col].sum().reset_index()
            traffic_brand = traffic_brand.sort_values(traffic_col, ascending=False).head(10)
            fig_t = px.bar(traffic_brand, x=bc_col, y=traffic_col, text_auto=".2s",
                           color_discrete_sequence=["#2563EB"])
            fig_t.update_layout(height=300, margin=dict(l=0, r=0, t=10, b=40),
                                 plot_bgcolor="white", paper_bgcolor="white",
                                 xaxis_tickangle=-30, font=dict(family="Inter"),
                                 yaxis=dict(showgrid=True, gridcolor="#F3F4F6"))
            st.plotly_chart(fig_t, use_container_width=True)

# ── ASIN Detail Table ──────────────────────────────────────────────────────────
st.markdown('<div class="section-header"><span>🔍 ASIN-Level Detail</span><span class="section-pill">FULL DATA</span></div>', unsafe_allow_html=True)

view_cols = [c for c in [col_map["prefix"], col_map["brand"], col_map["asin"],
                           col_map["net_price"], col_map["recommended"], col_map["analyst"],
                           col_map["sales"], col_map["traffic"]] if c and c in df.columns]
if view_cols:
    filter_col, _ = st.columns([2, 3])
    rec_filter = filter_col.selectbox("Filter by Recommendation", ["All", "Recommended", "Rejected", "Other"])
    view_df = df[view_cols + ["_rec_norm"]].copy()
    if rec_filter != "All":
        view_df = view_df[view_df["_rec_norm"] == rec_filter]
    st.dataframe(view_df.drop(columns=["_rec_norm"]), use_container_width=True, hide_index=True, height=320)
    st.caption(f"Showing {len(view_df):,} of {len(df):,} rows")
else:
    st.dataframe(df.head(200), use_container_width=True, hide_index=True)

# ── Executive Brief ────────────────────────────────────────────────────────────
st.markdown('<div class="section-header"><span>📝 Executive Brief</span><span class="section-pill">AUTO-GENERATED</span></div>', unsafe_allow_html=True)

top_analyst = a_grouped.iloc[0]["Analyst"] if not a_grouped.empty else "—"
top_analyst_audited = int(a_grouped.iloc[0]["Total Audited"]) if not a_grouped.empty else 0
top_analyst_rate = a_grouped.iloc[0]["Approval Rate %"] if not a_grouped.empty else 0
top_brand = brand_df.iloc[0][bc_col] if (not brand_df.empty and bc_col) else "—"
top_brand_count = int(brand_df.iloc[0]["Approved SKUs"]) if not brand_df.empty else 0

brief_lines = [
    f"• This week, **{kpis['total']:,} SKUs** were reviewed across the product catalogue ({week_label}).",
    f"• The team approved **{kpis['recommended']:,} items** (Approval Rate: **{kpis['rate']}%**) and rejected **{kpis['rejected']:,}** items.",
    f"• **{top_analyst}** was the top-performing analyst with **{top_analyst_audited} items audited** and an individual approval rate of **{top_analyst_rate:.1f}%**.",
    f"• **{top_brand}** leads brand performance with **{top_brand_count} approved SKUs** — recommended for priority sourcing and catalogue expansion.",
]
if not a_grouped.empty and len(a_grouped) > 1:
    others = ", ".join([f"{r['Analyst']} ({r['Approval Rate %']:.0f}%)" for _, r in a_grouped.iloc[1:].iterrows()])
    brief_lines.append(f"• Other analyst contributions: {others}.")
if not brand_df.empty and len(brand_df) > 1 and bc_col:
    runner_ups = ", ".join([f"**{r[bc_col]}** ({r['Approved SKUs']})" for _, r in brand_df.iloc[1:4].iterrows()])
    brief_lines.append(f"• Runner-up brands for consideration: {runner_ups}.")

brief_text = "\n".join(brief_lines)
st.markdown(
    f'<div style="background:#FAFBFF;border:1px solid #E5E7EB;border-radius:14px;padding:24px;'
    f'line-height:2;font-size:14px;color:#374151;">{brief_text}</div>',
    unsafe_allow_html=True,
)

# ── Export Section ─────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 Download Reports")
dl1, dl2, dl3 = st.columns(3)

summary_data = {
    "kpis": kpis,
    "analyst_df": a_grouped,
    "brand_df": brand_df,
    "bc_col": bc_col,
    "week_label": week_label,
    "report_date": str(report_date),
    "df_full": df,
    "col_map": col_map,
    "view_cols": view_cols,
}

with dl1:
    try:
        pptx_bytes = generate_pptx(summary_data)
        st.download_button(
            "📊 Download PowerPoint",
            data=pptx_bytes,
            file_name=f"VirVentures_AnalystReport_{report_date}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except Exception as e:
        st.error(f"PPTX error: {e}")

with dl2:
    try:
        xl_bytes = generate_excel(summary_data)
        st.download_button(
            "📋 Download Excel Report",
            data=xl_bytes,
            file_name=f"VirVentures_AnalystReport_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Excel error: {e}")

with dl3:
    csv_bytes = df[view_cols].to_csv(index=False).encode("utf-8") if view_cols else b""
    st.download_button(
        "📄 Download Filtered CSV",
        data=csv_bytes,
        file_name=f"VirVentures_FilteredData_{report_date}.csv",
        mime="text/csv",
    )

st.markdown("<br><br>", unsafe_allow_html=True)
