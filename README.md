# 📊 Vir Ventures — Weekly Analyst Performance Tracker

A premium Streamlit dashboard for weekly product audit tracking with executive-level
PowerPoint and Excel report exports.

---

## 🚀 Deploy to Streamlit Cloud (5 minutes)

### Step 1 — Push to GitHub
1. Create a new GitHub repository (e.g. `vir-ventures-tracker`)
2. Upload all files in this folder to the repository root

### Step 2 — Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **New app**
4. Select your repository, branch `main`, file `app.py`
5. Click **Deploy** → your public URL is ready in ~2 minutes

---

## 📁 File Structure

```
├── app.py               ← Main Streamlit app
├── export_pptx.py       ← PowerPoint report generator (6 executive slides)
├── export_excel.py      ← Excel report generator (4 formatted sheets)
├── requirements.txt     ← Python dependencies
├── .streamlit/
│   └── config.toml      ← Orange/white theme config
└── README.md
```

---

## 📋 Data Format

Upload your weekly Excel or CSV audit file. Key columns detected automatically:

| Column Name (any of these) | Purpose |
|---|---|
| `Recommended`, `AN`, `Rec` | Yes / Yes, low qty / No |
| `Analysist`, `Analyst`, `AR` | Analyst name (e.g. Babul, Chanchal) |
| `Brand`, `Q` | Brand/vendor name |
| `Output ASIN`, `ASIN`, `P` | Product ASIN |
| `Prefix` | Category prefix (BLFN, SPQT etc.) |
| `Sum of Net Price`, `Net Price` | Price per SKU |
| `Sales`, `Revenue` | Vendor sales data |
| `Traffic`, `Sessions` | Traffic data |

---

## 📊 Dashboard Features

### Executive KPIs
- Catalogue Size, Recommended, Rejected, Approval Rate

### Analyst Performance (Micro Tracking)
- Individual audit throughput table
- Grouped bar chart (Audited vs Recommended)
- Approval rate comparison with 70% target line
- Colour-coded signals (🟢 High / 🟡 Mid / 🔴 Low)

### Brand Intelligence
- Top brands by approved SKU count
- Horizontal bar chart

### Vendor Sales & Traffic
- Revenue by brand (if sales column present)
- Traffic/sessions by brand

### ASIN-Level Detail
- Filterable full data table

### Executive Brief
- Auto-generated bullet summary

---

## 📥 Exports

### PowerPoint (6 slides)
1. Cover slide — Navy + Orange executive theme
2. Executive Summary — KPI cards + highlights + bullet insights
3. Analyst Performance — Formatted table + score cards
4. Brand Analysis — Horizontal bar chart
5. SKU Recommendation Breakdown — Donut metrics + stacked bar
6. Closing / Next Steps

### Excel (4 sheets)
1. 📊 Dashboard — KPI summary + analyst + brand tables
2. 👥 Analyst Detail — Full analyst breakdown with ASIN list
3. 🏷️ Brand Intelligence — Ranked brand table + embedded bar chart
4. 📄 Raw Data — Full filtered data with colour-coded rows + auto-filter

---

## 🔧 Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open: http://localhost:8501
