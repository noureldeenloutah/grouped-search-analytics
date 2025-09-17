import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re, os, logging
from datetime import datetime
import pytz

# Optional packages
try:
    from st_aggrid import AgGrid, GridOptionsBuilder
    AGGRID_OK = True
except Exception:
    AGGRID_OK = False

try:
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    WORDCLOUD_OK = True
except Exception:
    WORDCLOUD_OK = False

# ----------------- Page config -----------------
st.set_page_config(page_title="🔥 Lady Care — Ultimate Search Analytics", layout="wide", page_icon="✨")
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ----------------- CSS / UI enhancements -----------------
st.markdown("""
<style>
/* Global styling */
body {
    font-family: 'Segoe UI', 'Arial', sans-serif;
    background: #F9FAFB;
}

/* Sidebar */
.sidebar .sidebar-content {
    background: linear-gradient(135deg, #FF5A6E 0%, #FFF7E8 100%);
    border-radius: 12px;
    padding: 15px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}
.sidebar .sidebar-content h1, .sidebar .sidebar-content * {
    color: #1A3C5E !important;
}

/* Header */
.main-header {
    font-size: 2.5rem;
    font-weight: 900;
    color: #FF5A6E;
    text-align: center;
    margin-bottom: 0.3rem;
    text-shadow: 1px 1px 3px rgba(0,0,0,0.1);
}

/* Subtitle */
.sub-header {
    font-size: 1.1rem;
    color: #0B486B;
    text-align: center;
    margin-bottom: 1rem;
}

/* Welcome section */
.welcome-box {
    background: linear-gradient(90deg, #FFEFEF, #E6F3FA);
    padding: 20px;
    border-radius: 12px;
    margin-bottom: 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    text-align: center;
}
.welcome-box h2 {
    color: #FF5A6E;
    font-size: 1.8rem;
    margin-bottom: 10px;
}
.welcome-box p {
    color: #0B486B;
    font-size: 1rem;
}

/* KPI card */
.kpi {
    background: linear-gradient(90deg, #FFFFFF, #F9FAFB);
    padding: 16px;
    border-radius: 12px;
    text-align: center;
    box-shadow: 0 6px 14px rgba(0,0,0,0.08);
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi:hover {
    transform: translateY(-6px);
    box-shadow: 0 8px 20px rgba(0,0,0,0.12);
}
.kpi .value {
    font-size: 1.8rem;
    font-weight: 800;
    color: #0B486B;
}
.kpi .label {
    color: #6C7A89;
    font-size: 0.95rem;
}

/* Insight box */
.insight-box {
    background: #FFF8F3;
    padding: 15px;
    border-left: 6px solid #FF8A7A;
    border-radius: 8px;
    margin-bottom: 15px;
    transition: transform 0.2s;
}
.insight-box:hover {
    transform: translateX(5px);
}
.insight-box h4 {
    margin: 0 0 8px 0;
    color: #0B486B;
}
.insight-box p {
    margin: 0;
    color: #444;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    background: #F0F8FF;
    padding: 10px;
    border-radius: 10px;
}
.stTabs [data-baseweb="tab"] {
    height: 50px;
    border-radius: 8px;
    padding: 12px;
    font-weight: 700;
    background: #FFFFFF;
    color: #0B486B;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(90deg, #FF5A6E, #FFB085);
    color: #FFFFFF !important;
}
.stTabs [data-baseweb="tab"]:hover {
    background: #FFEFEF;
    color: #FF5A6E;
}

/* Footer */
.footer {
    text-align: center;
    padding: 15px 0;
    color: #5F6B73;
    font-size: 0.9rem;
    margin-top: 20px;
    border-top: 2px solid #FF8A7A;
    background: linear-gradient(180deg, #F9FAFB, #FFFFFF);
}
.footer a {
    color: #FF5A6E;
    text-decoration: none;
}
.footer a:hover {
    text-decoration: underline;
}

/* Dataframe and AgGrid */
.dataframe, .stDataFrame {
    border-radius: 8px;
    overflow: hidden;
}
.stDataFrame table {
    background: #FFFFFF;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# ----------------- Helpers -----------------
def safe_read_excel(path):
    """Read Excel into dict of DataFrames (sheet_name -> df)."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    xls = pd.ExcelFile(path)
    sheets = {}
    for name in xls.sheet_names:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name)
        except Exception as e:
            logger.warning(f"Could not read sheet {name}: {e}")
    return sheets

def extract_keywords(text: str):
    """Extract words (Arabic & Latin & numbers) without correcting spelling."""
    if not isinstance(text, str):
        return []
    tokens = re.findall(r'[\u0600-\u06FF\w%+\-]+', text)
    return [t.strip().lower() for t in tokens if len(t.strip())>0]

def prepare_queries_df(df: pd.DataFrame):
    """Normalize columns, create derived metrics and time buckets."""
    df = df.copy()
    # Keep search text as-is (user requested not to correct spellings)
    if 'normalized_query' not in df.columns:
        if 'search' in df.columns:
            df['normalized_query'] = df['search'].astype(str)
        elif 'search_term' in df.columns:
            df['normalized_query'] = df['search_term'].astype(str)
        else:
            df['normalized_query'] = df.iloc[:,0].astype(str)

    # Ensure start_date/end_date exist and handle date parsing
    for c in ['start_date', 'end_date', 'Date', 'date']:
        if c in df.columns:
            # Check if the column is already datetime
            if pd.api.types.is_datetime64_any_dtype(df[c]):
                df['Date'] = df[c]
            else:
                # Check if the column is numeric (Excel serial dates)
                if pd.to_numeric(df[c], errors='coerce').notna().any():
                    df['Date'] = pd.to_datetime(df[c], errors='coerce', unit='D', origin='1899-12-30')
                else:
                    # Assume standard date strings (e.g., YYYY-MM-DD)
                    df['Date'] = pd.to_datetime(df[c], errors='coerce')
            break
    if 'Date' not in df.columns:
        df['Date'] = pd.NaT

    # Impressions
    imp_cols = [c for c in df.columns if c.lower() in ('total_impressions_3m','impressions','count','count_search','search_count')]
    if imp_cols:
        df['impressions'] = pd.to_numeric(df[imp_cols[0]], errors='coerce').fillna(0)
    else:
        df['impressions'] = 0

    # Clicks
    click_cols = [c for c in df.columns if c.lower() in ('total_clicks_3m','clicks','click_count')]
    if click_cols:
        df['clicks'] = pd.to_numeric(df[click_cols[0]], errors='coerce').fillna(0)
    else:
        df['clicks'] = 0

    # Conversions
    conv_cols = [c for c in df.columns if c.lower() in ('total_conversions_3m','conversions','conversion_count')]
    if conv_cols:
        df['conversions'] = pd.to_numeric(df[conv_cols[0]], errors='coerce').fillna(0)
    else:
        df['conversions'] = 0

    # Revenue if present
    rev_cols = [c for c in df.columns if c.lower().startswith('rev') or c.lower().startswith('revenue')]
    if rev_cols:
        df['revenue'] = pd.to_numeric(df[rev_cols[0]], errors='coerce').fillna(0)
    else:
        df['revenue'] = 0

    # Rates
    df['ctr'] = df.apply(lambda r: r['clicks']/r['impressions'] if r['impressions']>0 else 0, axis=1) * 100
    df['cr'] = df.apply(lambda r: r['conversions']/r['clicks'] if r['clicks']>0 else 0, axis=1) * 100
    df['classical_cr'] = df.apply(lambda r: r['conversions']/r['clicks'] if r['clicks']>0 else 0, axis=1) * 100

    # Time buckets
    df['year'] = df['Date'].dt.year
    df['month'] = df['Date'].dt.strftime('%b %Y')
    df['month_short'] = df['Date'].dt.strftime('%b')
    df['day_of_week'] = df['Date'].dt.day_name()

    # Text features
    df['query_length'] = df['normalized_query'].astype(str).apply(len)
    df['keywords'] = df['normalized_query'].apply(extract_keywords)

    # Arabic brand first word
    if 'Arabic description' in df.columns:
        df['brand_ar'] = df['Arabic description'].astype(str).str.split().str[0].fillna('').str.strip()
    else:
        df['brand_ar'] = ''

    # Canonical brand
    if 'item_brand' in df.columns:
        df['brand'] = df['item_brand']
    elif 'Brand' in df.columns:
        df['brand'] = df['Brand']
    elif 'brand' not in df.columns:
        df['brand'] = None

    # Category/subcategory/department
    for c in ['category','Category','category_main','Category Main']:
        if c in df.columns:
            df['category'] = df[c]
            break
    for c in ['sub_category','Sub Category','SubCategory']:
        if c in df.columns:
            df['sub_category'] = df[c]
            break
    for c in ['department','Department','DepartmentName']:
        if c in df.columns:
            df['department'] = df[c]
            break

    # Device / country
    if 'Device_Type' in df.columns:
        df['device'] = df['Device_Type']
    if 'Country' in df.columns:
        df['country'] = df['Country']

    return df

# ----------------- Data load (upload or default) -----------------
st.sidebar.title("📁 Upload Data")
upload = st.sidebar.file_uploader("Upload Excel (multi-sheet) or CSV (queries)", type=['xlsx','csv'])
if upload is not None:
    if upload.name.endswith('.xlsx'):
        try:
            sheets = pd.read_excel(upload, sheet_name=None)
        except Exception as e:
            st.error(f"Error reading uploaded Excel: {e}")
            st.stop()
    else:
        try:
            df_csv = pd.read_csv(upload)
            sheets = {'queries_clustered': df_csv}
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
            st.stop()
else:
    default_path = "Lady Care Preprocessed Data.xlsx"
    if os.path.exists(default_path):
        try:
            sheets = safe_read_excel(default_path)
        except Exception as e:
            st.error(f"Failed to load default Excel: {e}")
            st.stop()
    else:
        st.info("No file uploaded and default Excel not found. Please upload your preprocessed file (.xlsx or .csv).")
        st.stop()

# ----------------- Choose main queries sheet -----------------
sheet_keys = list(sheets.keys())
preferred = [k for k in ['queries_clustered','queries_dedup','queries','queries_clustered_preprocessed'] if k in sheets]
if preferred:
    main_key = preferred[0]
else:
    main_key = sheet_keys[0]

raw_queries = sheets[main_key]
queries = prepare_queries_df(raw_queries)

# Load additional summary sheets if present
brand_summary = sheets.get('brand_summary', None)
category_summary = sheets.get('category_summary', None)
subcategory_summary = sheets.get('subcategory_summary', None)
generic_type = sheets.get('generic_type', None)

# ----------------- Filters (no sampling) -----------------
st.sidebar.header("🔎 Filters")
min_date = queries['Date'].min()
max_date = queries['Date'].max()
if pd.isna(min_date):
    min_date = None
if pd.isna(max_date):
    max_date = None

date_range = st.sidebar.date_input("📅 Select Date Range", value=[min_date, max_date] if min_date is not None and max_date is not None else [])

if isinstance(date_range, (list, tuple)) and len(date_range) == 2 and date_range[0] is not None:
    start_date, end_date = date_range
    queries = queries[(queries['Date']>=pd.to_datetime(start_date)) & (queries['Date']<=pd.to_datetime(end_date))]

# Multi-select filters helper
def multi_filter(df, col, label, emoji):
    if col not in df.columns:
        return df
    opts = sorted(df[col].dropna().astype(str).unique().tolist())
    sel = st.sidebar.multiselect(f"{emoji} {label}", options=opts, default=opts)
    if not sel or len(sel)==len(opts):
        return df
    else:
        return df[df[col].astype(str).isin(sel)]

queries = multi_filter(queries, 'brand', 'Brand(s)', '🏷')
queries = multi_filter(queries, 'department', 'Department(s)', '🏬')
queries = multi_filter(queries, 'category', 'Category(ies)', '📦')
queries = multi_filter(queries, 'sub_category', 'Sub Category(ies)', '🧴')
queries = multi_filter(queries, 'country', 'Country(s)', '🌍')
queries = multi_filter(queries, 'device', 'Device Type(s)', '📱')

text_filter = st.sidebar.text_input("🔍 Filter queries by text (contains)")
if text_filter:
    queries = queries[queries['normalized_query'].str.contains(re.escape(text_filter), case=False, na=False)]

st.sidebar.markdown(f"**📊 Rows after filters:** {len(queries):,}")

# ----------------- Welcome Message -----------------
st.markdown("""
<div class="welcome-box">
    <h2>👋 Welcome to Lady Care Search Analytics! ✨</h2>
    <p>Explore search patterns, brand performance, and actionable insights. Use the sidebar to filter data, navigate tabs to dive deep, and download results for your reports!</p>
</div>
""", unsafe_allow_html=True)

# ----------------- KPI cards -----------------
st.markdown('<div class="main-header">🔥 Lady Care — Ultimate Search Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Uncover powerful insights from the <b>search</b> column with vibrant visuals and actionable pivots</div>', unsafe_allow_html=True)

total_impr = int(queries['impressions'].sum())
total_clicks = int(queries['clicks'].sum())
total_conv = int(queries['conversions'].sum())
overall_ctr = (queries['clicks'].sum()/queries['impressions'].sum()) * 100 if queries['impressions'].sum()>0 else 0
overall_cr = (queries['conversions'].sum()/queries['clicks'].sum()) * 100 if queries['clicks'].sum()>0 else 0
total_revenue = float(queries['revenue'].sum()) if 'revenue' in queries.columns else 0.0

c1,c2,c3,c4,c5 = st.columns(5)
with c1:
    st.markdown(f"<div class='kpi'><div class='value'>{total_impr:,}</div><div class='label'>✨ Total Impressions</div></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='kpi'><div class='value'>{total_clicks:,}</div><div class='label'>👆 Total Clicks</div></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='kpi'><div class='value'>{total_conv:,}</div><div class='label'>🎯 Total Conversions</div></div>", unsafe_allow_html=True)
with c4:
    st.markdown(f"<div class='kpi'><div class='value'>{overall_ctr:.2f}%</div><div class='label'>📈 Overall CTR</div></div>", unsafe_allow_html=True)
with c5:
    st.markdown(f"<div class='kpi'><div class='value'>{overall_cr:.2f}%</div><div class='label'>💡 Overall CR</div></div>", unsafe_allow_html=True)

# ----------------- Tabs -----------------
tab_overview, tab_search, tab_brand, tab_category, tab_subcat, tab_generic, tab_time, tab_pivot, tab_insights, tab_export = st.tabs([
    "📈 Overview","🔍 Search Analysis","🏷 Brand","📦 Category","🧴 Subcategory","🛠 Generic Type",
    "⏰ Time Analysis","📊 Pivot Builder","💡 Insights & Qs","⬇ Export"
])

# ----------------- Overview -----------------
with tab_overview:
    st.header("📈 Overview & Quick Wins")
    st.markdown("Quick visuals to spot trends and take immediate action. 🚀")
    colA, colB = st.columns([2,1])
    with colA:
        imp_trend = queries.groupby('Date')['impressions'].sum().reset_index()
        if not imp_trend.empty:
            fig = px.line(imp_trend, x='Date', y='impressions', title='Impressions over Time', labels={'impressions':'Impressions'}, color_discrete_sequence=px.colors.qualitative.D3)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No date-impression data to plot.")
    with colB:
        st.markdown("**Top 10 Queries (Impressions)**")
        top10 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(10)
        st.dataframe(top10.rename(columns={'normalized_query':'Query'}), use_container_width=True)

    st.markdown("---")
    st.subheader("🌍 Geography & Device Snapshot")
    g1,g2 = st.columns(2)
    with g1:
        if 'country' in queries.columns and queries['country'].notna().any():
            ct = queries.groupby('country')['impressions'].sum().reset_index().sort_values('impressions', ascending=False)
            st.plotly_chart(px.pie(ct, names='country', values='impressions', title='Impression Share by Country', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
        else:
            st.info("Country column not available.")
    with g2:
        if 'device' in queries.columns and queries['device'].notna().any():
            dv = queries.groupby('device')['impressions'].sum().reset_index().sort_values('impressions', ascending=False)
            st.plotly_chart(px.bar(dv, x='device', y='impressions', title='Impressions by Device', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        else:
            st.info("Device column not available.")

# ----------------- Search Analysis (core) -----------------
with tab_search:
    st.header("🔍 Search Column — Deep Dive")
    st.markdown("Analyze raw search queries (no spelling corrections) with keyword frequency, query length, and long-tail insights. 🎯")

    kw_series = queries['keywords'].explode().dropna()
    kw_counts = kw_series.value_counts().reset_index().rename(columns={'index':'keyword',0:'count'})
    if not kw_counts.empty:
        st.subheader("Top 30 Words in Search Queries")
        st.plotly_chart(px.bar(kw_counts.head(30).iloc[::-1], x='count', y='keyword', orientation='h', title='Top Words in Queries', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    else:
        st.info("No keywords extracted from queries.")

    if WORDCLOUD_OK and not kw_counts.empty:
        st.subheader("🌟 Word Cloud")
        freqs = dict(zip(kw_counts['keyword'], kw_counts['count']))
        wc = WordCloud(width=900, height=350, background_color='white', collocations=False, font_path=None).generate_from_frequencies(freqs)
        fig, ax = plt.subplots(figsize=(12,4))
        ax.imshow(wc, interpolation='bilinear')
        ax.axis('off')
        st.pyplot(fig)
    else:
        if not WORDCLOUD_OK:
            st.info("Install 'wordcloud' and 'matplotlib' to enable word cloud (pip install wordcloud matplotlib).")

    st.subheader("📏 Query Length & Performance")
    ql = queries.groupby('query_length').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
    ql['ctr'] = ql.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
    if not ql.empty:
        st.plotly_chart(px.scatter(ql, x='query_length', y='ctr', size='impressions', title='Query Length vs CTR (Size=Impressions)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
    else:
        st.info("No query length data.")

    st.subheader("📊 Long-Tail vs Short-Tail")
    queries['is_long_tail'] = queries['query_length'] >= 20
    lt = queries.groupby('is_long_tail').agg(impressions=('impressions','sum'), conversions=('conversions','sum')).reset_index()
    lt['label'] = lt['is_long_tail'].map({True:'Long-tail (>=20 chars)', False:'Short-tail'})
    if not lt.empty:
        st.plotly_chart(px.pie(lt, names='label', values='impressions', title='Impression Share: Long-Tail vs Short-Tail', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    else:
        st.info("No long-tail information.")

# ----------------- Brand Tab -----------------
with tab_brand:
    st.header("🏷 Brand Insights")
    st.markdown("Explore brand demand and performance metrics. 🚀")

    if brand_summary is not None:
        st.subheader("📋 Brand Summary (Sheet)")
        st.dataframe(brand_summary, use_container_width=True)

    if 'brand' in queries.columns and queries['brand'].notna().any():
        bs = queries.groupby('brand').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        bs['ctr'] = bs.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        bs['cr'] = bs.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        st.plotly_chart(px.bar(bs.sort_values('impressions', ascending=False).head(20), x='brand', y='impressions', title='Top Brands by Impressions', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        st.plotly_chart(px.scatter(bs, x='impressions', y='ctr', size='conversions', color='brand', title='Brand: Impressions vs CTR (Size=Conversions)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)

        st.subheader("🔑 Top Keywords per Brand")
        rows = []
        for brand, grp in queries.groupby('brand'):
            kw = Counter([w for sub in grp['keywords'] for w in sub])
            for k,cnt in kw.most_common(8):
                rows.append({'brand':brand,'keyword':k,'count':cnt})
        df_bkw = pd.DataFrame(rows)
        if not df_bkw.empty:
            pivot_bkw = df_bkw.pivot_table(index='brand', columns='keyword', values='count', fill_value=0)
            if AGGRID_OK:
                gb = GridOptionsBuilder.from_dataframe(pivot_bkw.reset_index())
                gb.configure_grid_options(enableRangeSelection=True, pagination=True)
                AgGrid(pivot_bkw.reset_index(), gridOptions=gb.build(), height=400)
            else:
                st.dataframe(pivot_bkw, use_container_width=True)
        else:
            st.info("Not enough keyword data per brand.")
    else:
        st.info("Brand column not available in the dataset.")

# ----------------- Category Tab -----------------
with tab_category:
    st.header("📦 Category Insights")
    st.markdown("Analyze category-level performance and search patterns. 🌟")

    if category_summary is not None:
        st.subheader("📋 Category Summary (Sheet)")
        st.dataframe(category_summary, use_container_width=True)

    if 'category' in queries.columns and queries['category'].notna().any():
        cs = queries.groupby('category').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        cs['ctr'] = cs.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        cs['cr'] = cs.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        st.plotly_chart(px.bar(cs.sort_values('impressions', ascending=False), x='category', y='impressions', title='Impressions by Category', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        st.plotly_chart(px.bar(cs.sort_values('cr', ascending=False), x='category', y='cr', title='Conversion Rate by Category (%)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)

        st.subheader("🔑 Top Keywords per Category")
        rows = []
        for cat, grp in queries.groupby('category'):
            kw = Counter([w for sub in grp['keywords'] for w in sub])
            for k,cnt in kw.most_common(8):
                rows.append({'category':cat,'keyword':k,'count':cnt})
        df_ckw = pd.DataFrame(rows)
        if not df_ckw.empty:
            pivot_ckw = df_ckw.pivot_table(index='category', columns='keyword', values='count', fill_value=0)
            if AGGRID_OK:
                gb = GridOptionsBuilder.from_dataframe(pivot_ckw.reset_index())
                gb.configure_grid_options(enableRangeSelection=True, pagination=True)
                AgGrid(pivot_ckw.reset_index(), gridOptions=gb.build(), height=400)
            else:
                st.dataframe(pivot_ckw, use_container_width=True)
        else:
            st.info("Not enough keyword data per category.")
    else:
        st.info("No category column found.")

# ----------------- Subcategory Tab -----------------
with tab_subcat:
    st.header("🧴 Subcategory Insights")
    st.markdown("Dive into subcategory performance and search trends. 🚀")

    if subcategory_summary is not None:
        st.subheader("📋 Subcategory Summary (Sheet)")
        st.dataframe(subcategory_summary, use_container_width=True)

    if 'sub_category' in queries.columns and queries['sub_category'].notna().any():
        sc = queries.groupby('sub_category').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        sc['ctr'] = sc.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        st.plotly_chart(px.bar(sc.sort_values('impressions', ascending=False).head(30), x='sub_category', y='impressions', title='Top Subcategories by Impressions', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    else:
        st.info("No sub_category column present in dataset.")

# ----------------- Generic Type Tab -----------------
with tab_generic:
    st.header("🛠 Generic Type Insights")
    st.markdown("Explore generic type performance (if available). 🌟")

    if generic_type is not None:
        st.subheader("📋 Generic Type Summary (Sheet)")
        st.dataframe(generic_type, use_container_width=True)
    else:
        st.info("No 'generic_type' sheet provided.")

# ----------------- Time Analysis Tab -----------------
with tab_time:
    st.header("⏰ Temporal Analysis & Seasonality")
    st.markdown("Uncover monthly and weekly trends to optimize campaigns. 📅")

    if queries['month'].notna().any():
        monthly = queries.groupby('month').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        monthly['ctr'] = monthly.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        try:
            monthly['month_dt'] = pd.to_datetime(monthly['month'], format='%b %Y', errors='coerce')
            monthly = monthly.sort_values('month_dt')
        except:
            monthly = monthly.sort_values('month')
        st.plotly_chart(px.line(monthly, x='month', y='impressions', title='Monthly Impressions', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        st.plotly_chart(px.line(monthly, x='month', y='ctr', title='Monthly Average CTR (%)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
    else:
        st.info("No month data to plot.")

    if queries['day_of_week'].notna().any():
        dow = queries.groupby('day_of_week').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
        days_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
        dow['day_of_week'] = pd.Categorical(dow['day_of_week'], categories=days_order, ordered=True)
        dow = dow.sort_values('day_of_week')
        dow['ctr'] = dow.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        st.plotly_chart(px.bar(dow, x='day_of_week', y='ctr', title='CTR by Day of Week (%)', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    else:
        st.info("No day_of_week column in data.")

# ----------------- Pivot Builder Tab -----------------
with tab_pivot:
    st.header("📊 Pivot Builder & Prebuilt Pivots")
    st.markdown("Create custom pivots or explore prebuilt tables for quick insights. 🔧")

    st.subheader("📋 Prebuilt: Brand × Query (Top 300)")
    if 'brand' in queries.columns:
        pv = queries.groupby(['brand','normalized_query']).agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        pv['ctr'] = pv.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        pv_top = pv.sort_values('impressions', ascending=False).head(300)
        if AGGRID_OK:
            gb = GridOptionsBuilder.from_dataframe(pv_top)
            gb.configure_grid_options(enableRangeSelection=True, pagination=True)
            AgGrid(pv_top, gridOptions=gb.build(), height=400)
        else:
            st.dataframe(pv_top.head(100), use_container_width=True)
    else:
        st.info("Brand column missing for this pivot.")

    st.subheader("📋 Prebuilt: Month × Subcategory (Clicks)")
    if 'month' in queries.columns and 'sub_category' in queries.columns:
        pv2 = queries.groupby(['month','sub_category']).agg(clicks=('clicks','sum')).reset_index()
        pv2_pivot = pv2.pivot(index='month', columns='sub_category', values='clicks').fillna(0)
        if AGGRID_OK:
            gb = GridOptionsBuilder.from_dataframe(pv2_pivot.reset_index())
            gb.configure_grid_options(enableRangeSelection=True, pagination=True)
            AgGrid(pv2_pivot.reset_index(), gridOptions=gb.build(), height=400)
        else:
            st.dataframe(pv2_pivot, use_container_width=True)
    else:
        st.info("Month or sub_category missing for this pivot.")

    st.markdown("---")
    st.subheader("🔧 Custom Pivot Builder")
    columns = queries.columns.tolist()
    idx = st.multiselect("Rows (Index)", options=columns, default=['normalized_query'])
    cols = st.multiselect("Columns", options=columns, default=['brand'])
    val = st.selectbox("Value (Measure)", options=['impressions','clicks','conversions','revenue'], index=0)
    aggfunc = st.selectbox("Aggregation", options=['sum','mean','count'], index=0)
    if st.button("Generate Pivot"):
        try:
            pivot = pd.pivot_table(queries, values=val, index=idx, columns=cols, aggfunc=aggfunc, fill_value=0)
            if AGGRID_OK:
                gb = GridOptionsBuilder.from_dataframe(pivot.reset_index())
                gb.configure_grid_options(enableRangeSelection=True, pagination=True)
                AgGrid(pivot.reset_index(), gridOptions=gb.build(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
            st.download_button("⬇ Download Pivot CSV", pivot.to_csv().encode('utf-8'), file_name='custom_pivot.csv')
        except Exception as e:
            st.error(f"Pivot generation error: {e}")

# ----------------- Insights & Questions -----------------
with tab_insights:
    st.header("💡 Insights & Actionable Questions (40+)")
    st.markdown("Actionable insights focused on the **search** column, with pivot tables and visuals. 🚀")

    def q_expand(title, explanation, render_fn, icon="💡"):
        with st.expander(f"{icon} {title}", expanded=False):
            st.markdown(f"<div class='insight-box'><h4>Why & How to Use</h4><p>{explanation}</p></div>", unsafe_allow_html=True)
            try:
                render_fn()
            except Exception as e:
                st.error(f"Rendering error: {e}")

    # Q1: Top queries by impressions
    def q1():
        out = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q1 — Top Queries by Impressions (Top 30)", "Which queries drive the most traffic? Prioritize for search tuning and inventory.", q1, "📈")

    # Q2: High impressions, low CTR
    def q2():
        df2 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
        df2['ctr'] = df2.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        out = df2[(df2['impressions']>=df2['impressions'].quantile(0.6)) & (df2['ctr']<=df2['ctr'].quantile(0.3))].sort_values('impressions', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q2 — High Impressions, Low CTR", "Queries with high traffic but low engagement. Improve relevance, snippets, or imagery.", q2, "⚠️")

    # Q3: High CTR, low revenue
    def q3():
        if 'revenue' not in queries.columns:
            st.info("Revenue not available.")
            return
        df3 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), revenue=('revenue','sum')).reset_index()
        df3['ctr'] = df3.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        out = df3[(df3['ctr']>=df3['ctr'].quantile(0.75)) & (df3['revenue']<=df3['revenue'].quantile(0.4))].sort_values('ctr', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q3 — High CTR, Low Revenue", "Good engagement but low monetization. Check pricing, stock, or funnel issues.", q3, "💸")

    # Q4: Top queries by conversion rate
    def q4():
        df4 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df4 = df4[df4['impressions']>=50]
        df4['cr'] = df4.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df4.sort_values('cr', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q4 — Top Queries by Conversion Rate (Min Impressions=50)", "High-converting queries for paid promotions.", q4, "🎯")

    # Q5: Long-tail contribution
    def q5():
        lt = queries[queries['query_length']>=20]
        st.markdown(f"Long-tail rows: {len(lt):,} / total {len(queries):,}")
        st.plotly_chart(px.pie(names=['Long-tail','Rest'], values=[lt['impressions'].sum(), queries['impressions'].sum()-lt['impressions'].sum()], title='Impression Share: Long-Tail vs Rest', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    q_expand("Q5 — Long-Tail vs Short-Tail (>=20 chars)", "How much traffic comes from long-tail queries? Key for content strategy.", q5, "📏")

    # Q6: Misspelling types
    if 'misspelling_type' in queries.columns:
        def q6():
            ms = queries.groupby('misspelling_type').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
            st.plotly_chart(px.bar(ms, x='misspelling_type', y='impressions', title='Impressions by Misspelling Type', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
        q_expand("Q6 — Misspelling Types Impact", "Which misspelling categories drive traffic?", q6, "🔍")

    # Q7: Brand vs generic share (pivot)
    def q7():
        if 'brand' in queries.columns:
            branded = queries[queries['brand'].notna() & (queries['brand']!='')]
            branded_share = branded['impressions'].sum()
            total = queries['impressions'].sum()
            st.markdown(f"Branded impressions: {branded_share:,} / Total: {total:,}  —  Share: {branded_share/total:.2%}")
            st.plotly_chart(px.pie(names=['Branded','Generic'], values=[branded_share, total-branded_share], title='Branded vs Generic Impression Share', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
            pivot = queries.pivot_table(values=['impressions','clicks','conversions'], index=['brand'], aggfunc='sum').reset_index()
            if AGGRID_OK:
                AgGrid(pivot, height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Brand column not present.")
    q_expand("Q7 — Branded vs Generic Queries (Pivot)", "Assess brand vs generic search intent with a pivot table.", q7, "🏷")

    # Q8: Rising queries MoM
    def q8():
        mom = queries.groupby(['month','normalized_query']).agg(impressions=('impressions','sum')).reset_index()
        if len(mom['month'].unique())<2:
            st.info("Not enough months to compute MoM.")
            return
        pivot = mom.pivot(index='normalized_query', columns='month', values='impressions').fillna(0)
        months_sorted = sorted(pivot.columns, key=lambda x: pd.to_datetime(x, format='%b %Y', errors='coerce') if isinstance(x,str) else x)
        if len(months_sorted)>=2:
            recent, prev = months_sorted[-1], months_sorted[-2]
            pivot['pct_change'] = (pivot[recent]-pivot[prev])/(pivot[prev].replace(0,np.nan))
            out = pivot.sort_values('pct_change', ascending=False).head(30).reset_index()
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Not enough months.")
    q_expand("Q8 — Top Rising Queries Month-over-Month", "Detect emerging demand for seasonal campaigns.", q8, "📈")

    # Q9: Zero-results queries
    if 'zero_results' in queries.columns:
        def q9():
            zr = queries[queries['zero_results']==True].groupby('normalized_query').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions',ascending=False)
            if AGGRID_OK:
                AgGrid(zr.head(50), height=400)
            else:
                st.dataframe(zr.head(50), use_container_width=True)
        q_expand("Q9 — Queries Returning Zero Results", "Fix indexing/product mapping for these queries.", q9, "⚠️")

    # Q10: Device split
    def q10():
        if 'device' in queries.columns:
            dev = queries.groupby(['device']).agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
            dev['ctr'] = dev.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
            dev['cr'] = dev.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
            if AGGRID_OK:
                AgGrid(dev, height=400)
            else:
                st.dataframe(dev, use_container_width=True)
        else:
            st.info("Device column missing.")
    q_expand("Q10 — Device Performance (Mobile/Desktop/Tablet)", "Check for mobile UX issues.", q10, "📱")

    # Q11: Query funnel snapshot
    def q11():
        snap = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(200)
        if AGGRID_OK:
            AgGrid(snap, height=400)
        else:
            st.dataframe(snap.head(100), use_container_width=True)
    q_expand("Q11 — Query Funnel Snapshot (Top 200)", "View top queries' funnel: impressions → clicks → conversions.", q11, "📊")

    # Q12: High revenue-per-conversion queries
    def q12():
        if 'revenue' not in queries.columns or queries['revenue'].sum()==0:
            st.info("Revenue not available.")
            return
        rr = queries.groupby('normalized_query').agg(revenue=('revenue','sum'), conversions=('conversions','sum')).reset_index()
        rr = rr[rr['conversions']>0]
        rr['rev_per_conv'] = rr['revenue']/rr['conversions']
        out = rr.sort_values('rev_per_conv', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q12 — High Revenue-per-Conversion Queries", "Identify high-LTV queries for promotion.", q12, "💸")

    # Q13: Low match confidence
    if 'match_confidence' in queries.columns:
        def q13():
            mc = queries.groupby('normalized_query').agg(mean_conf=('match_confidence','mean'), impressions=('impressions','sum')).reset_index().sort_values('mean_conf')
            if AGGRID_OK:
                AgGrid(mc.head(50), height=400)
            else:
                st.dataframe(mc.head(50), use_container_width=True)
        q_expand("Q13 — Low Match Confidence Queries", "Improve matching logic for low-confidence queries.", q13, "🔍")

    # Q14: Traffic concentration
    def q14():
        qq = queries.groupby('normalized_query').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions', ascending=False)
        top5n = max(1, int(0.05*len(qq)))
        share = qq.head(top5n)['impressions'].sum() / qq['impressions'].sum() if qq['impressions'].sum()>0 else 0
        st.markdown(f"Top 5% queries contribute **{share:.2%}** of impressions (top {top5n} queries).")
        st.plotly_chart(px.pie(names=['Top 5% Queries','Rest'], values=[qq.head(top5n)['impressions'].sum(), qq['impressions'].sum()-qq.head(top5n)['impressions'].sum()], title='Traffic Concentration: Top 5% Queries', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    q_expand("Q14 — Traffic Concentration (Top 5%)", "Prioritize top queries driving most traffic.", q14, "📈")

    # Q15: Keyword co-occurrence
    def q15():
        from itertools import combinations
        kw_lists = queries['keywords'].dropna().tolist()
        co = Counter()
        for kws in kw_lists:
            uniq = list(dict.fromkeys(kws))
            for a,b in combinations(uniq,2):
                co[(a,b)] += 1
        co_df = pd.DataFrame([{'kw_pair':f"{a} | {b}", 'count':c} for (a,b),c in co.items()]).sort_values('count', ascending=False).head(50)
        if not co_df.empty:
            if AGGRID_OK:
                AgGrid(co_df, height=400)
            else:
                st.dataframe(co_df, use_container_width=True)
        else:
            st.info("Not enough keyword co-occurrence data.")
    q_expand("Q15 — Keyword Co-Occurrence (Cross-Sell Proxy)", "Find keywords searched together for cross-sell opportunities.", q15, "🔗")

    # Q16: High searches, zero conversions
    def q16():
        dfm = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        out = dfm[(dfm['impressions']>=dfm['impressions'].quantile(0.7)) & (dfm['conversions']==0)].sort_values('impressions', ascending=False).head(40)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q16 — High Search Volume, Zero Conversions", "Fix product discovery or pricing for these queries.", q16, "⚠️")

    # Q17: Queries with many variants
    def q17():
        token_map = {}
        for q in queries['normalized_query'].dropna().unique():
            key = " ".join(extract_keywords(q)[:2])
            token_map.setdefault(key,0)
            token_map[key]+=1
        tok_df = pd.DataFrame([{'prefix':k,'count':v} for k,v in token_map.items()]).sort_values('count', ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(tok_df, height=400)
        else:
            st.dataframe(tok_df, use_container_width=True)
    q_expand("Q17 — Queries with Many Variants (Prefix Clustering)", "Identify queries with variants/typos for canonicalization.", q17, "🔍")

    # Q18: Day-part conversions
    if 'Date' in queries.columns and queries['Date'].notna().any():
        def q18():
            queries['hour'] = queries['Date'].dt.hour
            hr = queries.groupby('hour').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
            hr['ctr'] = hr.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
            st.plotly_chart(px.line(hr, x='hour', y='ctr', title='CTR by Hour of Day (%)', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        q_expand("Q18 — Time of Day Patterns (CTR by Hour)", "Identify peak hours for promotions.", q18, "⏰")

    # Q19: Top queries by CTR
    def q19():
        df19 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
        df19 = df19[df19['impressions']>=30]
        df19['ctr'] = df19.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        out = df19.sort_values('ctr', ascending=False).head(40)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q19 — Top Queries by CTR (Min Impressions=30)", "High-engagement queries for ad campaigns.", q19, "📈")

    # Q20: Low CTR & CR queries
    def q20():
        df20 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df20['ctr'] = df20.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        df20['cr'] = df20.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df20[(df20['impressions']>=df20['impressions'].quantile(0.6)) & (df20['ctr']<=df20['ctr'].quantile(0.25)) & (df20['cr']<=df20['cr'].quantile(0.25))].sort_values('impressions', ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q20 — High Impressions, Low CTR & CR", "Optimize search results for these underperforming queries.", q20, "⚠️")

    # Q21: Top keywords per category (pivot)
    def q21():
        if 'category' in queries.columns:
            rows = []
            for cat,grp in queries.groupby('category'):
                kw = Counter([w for sub in grp['keywords'] for w in sub])
                for k,cnt in kw.most_common(5):
                    rows.append({'category':cat,'keyword':k,'count':cnt})
            df21 = pd.DataFrame(rows)
            pivot = df21.pivot_table(index='category', columns='keyword', values='count', fill_value=0)
            if AGGRID_OK:
                AgGrid(pivot.reset_index(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Category not available.")
    q_expand("Q21 — Top Keywords per Category (Pivot)", "Understand category-specific search language for taxonomy.", q21, "📦")

    # Q22: Promotion-related queries
    def q22():
        promo_words = ['عرض','تخفيض','خصم','offer','sale','discount']
        mask = queries['normalized_query'].apply(lambda x: any(w in str(x).lower() for w in promo_words))
        if mask.any():
            promo_df = queries[mask].groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(promo_df, height=400)
            else:
                st.dataframe(promo_df, use_container_width=True)
        else:
            st.info("No promo-related queries found.")
    q_expand("Q22 — Queries with Promotion Words", "Leverage for promotional campaigns.", q22, "🎉")

    # Q23: Brand-inclusive queries
    def q23():
        if 'brand' in queries.columns:
            labeled = queries[queries['brand'].notna() & (queries['brand']!='')]
            brand_q = labeled.groupby('normalized_query').size().reset_index(name='count').sort_values('count', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(brand_q, height=400)
            else:
                st.dataframe(brand_q, use_container_width=True)
        else:
            st.info("Brand column missing.")
    q_expand("Q23 — Brand-Inclusive Queries", "High purchase intent queries with brands.", q23, "🏷")

    # Q24: Purchase-intent queries
    def q24():
        buy_words = ['شراء','اشتري','buy','سعر','price']
        mask = queries['normalized_query'].apply(lambda x: any(w in str(x).lower() for w in buy_words))
        if mask.any():
            buy_df = queries[mask].groupby('normalized_query').agg(impressions=('impressions','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(buy_df, height=400)
            else:
                st.dataframe(buy_df, use_container_width=True)
        else:
            st.info("No 'buy' intent queries detected.")
    q_expand("Q24 — Purchase-Intent Queries (Buy/Price)", "Prioritize for conversion flow optimization.", q24, "💸")

    # Q25: Queries with many results
    if 'results_count' in queries.columns:
        def q25():
            rc = queries.groupby('normalized_query').agg(results_mean=('results_count','mean'), impressions=('impressions','sum')).reset_index().sort_values('results_mean', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(rc, height=400)
            else:
                st.dataframe(rc, use_container_width=True)
        q_expand("Q25 — Queries Returning Many Results", "Check for UX improvements where too many results overwhelm.", q25, "🔍")

    # Q26: Top queries by conversions
    def q26():
        out = queries.groupby('normalized_query').agg(conversions=('conversions','sum'), impressions=('impressions','sum')).reset_index().sort_values('conversions', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q26 — Top Queries by Conversions", "Direct revenue drivers for inventory and bids.", q26, "🎯")

    # Q27: Month-by-month query trends (pivot)
    def q27():
        mom = queries.groupby(['month','normalized_query']).agg(impressions=('impressions','sum')).reset_index()
        topq = mom.groupby('normalized_query')['impressions'].sum().reset_index().sort_values('impressions', ascending=False).head(500)['normalized_query']
        sample = mom[mom['normalized_query'].isin(topq)].pivot(index='normalized_query', columns='month', values='impressions').fillna(0)
        if sample.shape[1] >= 2:
            if AGGRID_OK:
                AgGrid(sample.head(200).reset_index(), height=400)
            else:
                st.dataframe(sample.head(200), use_container_width=True)
        else:
            st.info("Not enough months to show seasonality.")
    q_expand("Q27 — Month-by-Month Query Trends (Pivot)", "Identify seasonal trends for campaign planning.", q27, "📅")

    # Q28: Device success metrics
    def q28():
        if 'device' in queries.columns:
            dev = queries.groupby(['device']).agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
            dev['ctr'] = dev.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
            dev['cr'] = dev.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
            if AGGRID_OK:
                AgGrid(dev, height=400)
            else:
                st.dataframe(dev, use_container_width=True)
        else:
            st.info("Device column missing.")
    q_expand("Q28 — Device-Level Success Metrics", "Compare mobile vs desktop performance.", q28, "📱")

    # Q29: Unique keywords by category
    def q29():
        if 'category' in queries.columns:
            uniq = queries.groupby('category').agg(unique_keywords=('keywords', lambda s: len(set([k for sub in s for k in sub])))).reset_index().sort_values('unique_keywords', ascending=False)
            if AGGRID_OK:
                AgGrid(uniq, height=400)
            else:
                st.dataframe(uniq, use_container_width=True)
        else:
            st.info("Category missing.")
    q_expand("Q29 — Unique Keywords by Category", "Measure search diversity for facet planning.", q29, "📦")

    # Q30: Long-term growth queries
    def q30():
        if queries['month'].nunique() < 2:
            st.info("Not enough months.")
            return
        months_sorted = sorted(queries['month'].dropna().unique(), key=lambda x: pd.to_datetime(x, format='%b %Y', errors='coerce'))
        first, last = months_sorted[0], months_sorted[-1]
        m1 = queries[queries['month']==first].groupby('normalized_query').agg(impressions=('impressions','sum')).rename(columns={'impressions':'m1'})
        m2 = queries[queries['month']==last].groupby('normalized_query').agg(impressions=('impressions','sum')).rename(columns={'impressions':'m2'})
        comp = m1.join(m2, how='outer').fillna(0)
        comp['pct_change'] = (comp['m2'] - comp['m1']) / comp['m1'].replace(0, np.nan)
        out = comp.sort_values('pct_change', ascending=False).head(50).reset_index()
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q30 — Long-Term Growth: First vs Last Month", "Find queries with significant growth or decline.", q30, "📈")

    # Q31: Top 50 queries (quick)
    def q31():
        out = queries.groupby('normalized_query').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions',ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q31 — Top 50 Queries (Quick)", "Quick ranking of top queries by impressions.", q31, "📊")

    # Q32: Top brands quick view
    def q32():
        if 'brand' in queries.columns:
            out = queries.groupby('brand').agg(impressions=('impressions','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Brand missing.")
    q_expand("Q32 — Top Brands (Quick)", "Quick brand ranking by impressions.", q32, "🏷")

    # Q33: Top subcategories quick view
    def q33():
        if 'sub_category' in queries.columns:
            out = queries.groupby('sub_category').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Subcategory missing.")
    q_expand("Q33 — Top Subcategories (Quick)", "Quick subcategory ranking.", q33, "🧴")

    # Q34: Top queries by revenue
    def q34():
        if 'revenue' in queries.columns:
            out = queries.groupby('normalized_query').agg(revenue=('revenue','sum')).reset_index().sort_values('revenue', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Revenue not available.")
    q_expand("Q34 — Top Queries by Revenue", "Direct revenue drivers.", q34, "💸")

    # Q35: Monthly impressions table
    def q35():
        out = queries.groupby('month').agg(impressions=('impressions','sum')).reset_index().sort_values('month')
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q35 — Monthly Impressions Table", "Month-level volumes for reporting.", q35, "📅")

    # Q36: Arabic brand first-word counts
    def q36():
        if 'brand_ar' in queries.columns and queries['brand_ar'].notna().any():
            out = queries.groupby('brand_ar').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Arabic brand-first not available.")
    q_expand("Q36 — Arabic Brand First-Word Counts", "Check Arabic brand term distribution.", q36, "🌍")

    # Q37: Top queries by clicks
    def q37():
        out = queries.groupby('normalized_query').agg(clicks=('clicks','sum'), impressions=('impressions','sum')).reset_index().sort_values('clicks', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q37 — Top Queries by Clicks", "High-engagement queries for ad optimization.", q37, "👆")

    # Q38: Category vs brand performance (pivot)
    def q38():
        if 'category' in queries.columns and 'brand' in queries.columns:
            pivot = queries.pivot_table(values=['impressions','clicks','conversions'], index=['category'], columns=['brand'], aggfunc='sum').fillna(0)
            if AGGRID_OK:
                AgGrid(pivot.reset_index(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Category or brand missing.")
    q_expand("Q38 — Category vs Brand Performance (Pivot)", "Analyze brand performance within categories.", q38, "📦🏷")

    # Q39: High impressions, low clicks by category
    def q39():
        if 'category' in queries.columns:
            df39 = queries.groupby(['category','normalized_query']).agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
            df39['ctr'] = df39.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
            out = df39[(df39['impressions']>=df39['impressions'].quantile(0.6)) & (df39['ctr']<=df39['ctr'].quantile(0.3))].sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Category missing.")
    q_expand("Q39 — High Impressions, Low Clicks by Category", "Identify category-specific queries needing optimization.", q39, "⚠️")

    # Q40: Top keywords by device (pivot)
    def q40():
        if 'device' in queries.columns:
            rows = []
            for dev,grp in queries.groupby('device'):
                kw = Counter([w for sub in grp['keywords'] for w in sub])
                for k,cnt in kw.most_common(5):
                    rows.append({'device':dev,'keyword':k,'count':cnt})
            df40 = pd.DataFrame(rows)
            pivot = df40.pivot_table(index='device', columns='keyword', values='count', fill_value=0)
            if AGGRID_OK:
                AgGrid(pivot.reset_index(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Device missing.")
    q_expand("Q40 — Top Keywords by Device (Pivot)", "Understand device-specific search behavior.", q40, "📱")

    st.info("Want more advanced questions (e.g., anomaly detection, semantic clustering)? I can add them with additional packages like scikit-learn or prophet.")

# ----------------- Export / Downloads -----------------
with tab_export:
    st.header("⬇ Export & Save")
    st.markdown("Download filtered data or sheets for reporting. 📥")
    st.download_button("Download Filtered Queries CSV", queries.to_csv(index=False).encode('utf-8'), file_name='filtered_queries.csv')
    for name, df_s in sheets.items():
        try:
            st.download_button(f"Download Sheet: {name}", df_s.to_csv(index=False).encode('utf-8'), file_name=f"{name}.csv", key=f"dl_{name}")
        except Exception:
            pass
    st.markdown("---")
    st.markdown("💡 Tip: Use the Pivot Builder tab to create custom tables and download them as CSV.")

# ----------------- Footer -----------------
st.markdown(f"""
<div class="footer">
✨ Lady Care Search Analytics — Noureldeen Mohamed""")