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
/* Mini Metric Card */
.mini-metric {
    background: linear-gradient(90deg, #FF5A6E, #FFB085);
    padding: 12px;
    border-radius: 10px;
    text-align: center;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    transition: transform 0.2s, box-shadow 0.2s;
    height: 100px; /* Fixed height for uniformity */
    display: flex;
    flex-direction: column;
    justify-content: center;
}
.mini-metric:hover {
    transform: translateY(-4px);
    box-shadow: 0 6px 16px rgba(0,0,0,0.12);
}
.mini-metric .value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #FFFFFF;
    margin-bottom: 4px;
}
.mini-metric .label {
    font-size: 0.9rem;
    color: #F9FAFB;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.mini-metric .icon {
    font-size: 1.2rem;
    color: #FFFFFF;
    margin-bottom: 6px;
    display: block;
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
    if not sheets:
        raise ValueError("No valid sheets found in the Excel file.")
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
    
    # Query text
    if 'search' in df.columns:
        df['normalized_query'] = df['search'].astype(str)
    else:
        df['normalized_query'] = df.iloc[:,0].astype(str)

    # Date
    if 'start_date' in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df['start_date']):
            df['Date'] = df['start_date']
        else:
            df['Date'] = pd.to_datetime(df['start_date'], unit='D', origin='1899-12-30', errors='coerce')
    else:
        df['Date'] = pd.NaT

    # Counts (renamed from impressions)
    if 'total_impressions over 3m' in df.columns:
        df['Counts'] = pd.to_numeric(df['total_impressions over 3m'], errors='coerce').fillna(0)
    else:
        df['Counts'] = 0

    # Clicks
    if 'count' in df.columns:
        df['clicks'] = pd.to_numeric(df['count'], errors='coerce').fillna(0)
    else:
        df['clicks'] = 0

    # Conversions (derived from clicks * conversion rate)
    if 'Converion Rate' in df.columns and 'count' in df.columns:
        df['conversions'] = (df['count'] * df['Converion Rate']).round().astype(int)
    else:
        df['conversions'] = 0

    # CTR and CR
    if 'Click Through Rate' in df.columns:
        df['ctr'] = pd.to_numeric(df['Click Through Rate'], errors='coerce').fillna(0) * 100
    else:
        df['ctr'] = df.apply(lambda r: r['clicks']/r['Counts'] if r['Counts']>0 else 0, axis=1) * 100
    
    if 'Converion Rate' in df.columns:
        df['cr'] = pd.to_numeric(df['Converion Rate'], errors='coerce').fillna(0) * 100
    else:
        df['cr'] = df.apply(lambda r: r['conversions']/r['clicks'] if r['clicks']>0 else 0, axis=1) * 100

    if 'classical_cr' in df.columns:
        df['classical_cr'] = pd.to_numeric(df['classical_cr'], errors='coerce').fillna(0) * 100
    else:
        df['classical_cr'] = df['cr']

    # No revenue in provided columns
    df['revenue'] = 0

    # Time buckets
    df['year'] = df['Date'].dt.year
    df['month'] = df['Date'].dt.strftime('%b %Y')
    df['month_short'] = df['Date'].dt.strftime('%b')
    df['day_of_week'] = df['Date'].dt.day_name()

    # Text features
    df['query_length'] = df['normalized_query'].astype(str).apply(len)
    df['keywords'] = df['normalized_query'].apply(extract_keywords)

    # No Arabic description in provided columns
    df['brand_ar'] = ''

    # Brand, Category, Subcategory, Department
    if 'Brand' in df.columns:
        df['brand'] = df['Brand']
    else:
        df['brand'] = None

    if 'Category' in df.columns:
        df['category'] = df['Category']
    else:
        df['category'] = None

    if 'Sub Category' in df.columns:
        df['sub_category'] = df['Sub Category']
    else:
        df['sub_category'] = None

    if 'Department' in df.columns:
        df['department'] = df['Department']
    else:
        df['department'] = None

    # Additional columns
    if 'underperforming' in df.columns:
        df['underperforming'] = df['underperforming']
    if 'averageClickPosition' in df.columns:
        df['average_click_position'] = df['averageClickPosition']
    if 'cluster_id' in df.columns:
        df['cluster_id'] = df['cluster_id']

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
            st.error(f"Failed to load default Excel: {e}. Please ensure the file exists and is a valid Excel file.")
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
try:
    queries = prepare_queries_df(raw_queries)
except Exception as e:
    st.error(f"Error processing queries sheet: {e}")
    st.stop()

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

total_counts = int(queries['Counts'].sum())
total_clicks = int(queries['clicks'].sum())
total_conv = int(queries['conversions'].sum())
overall_ctr = (queries['clicks'].sum()/queries['Counts'].sum()) * 100 if queries['Counts'].sum()>0 else 0
overall_cr = (queries['conversions'].sum()/queries['clicks'].sum()) * 100 if queries['clicks'].sum()>0 else 0
total_revenue = 0.0  # No revenue column

c1,c2,c3,c4,c5 = st.columns(5)
with c1:
    st.markdown(f"<div class='kpi'><div class='value'>{total_counts:,}</div><div class='label'>✨ Total Counts</div></div>", unsafe_allow_html=True)
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
    st.markdown("Quick visuals to spot trends and take immediate action. 🚀 Based on **queries_clustered** data (e.g., 17M+ Counts across categories).")

    # Accuracy Fix: Ensure Date conversion (Excel serial)
    if not queries['Date'].dtype == 'datetime64[ns]':
        queries['Date'] = pd.to_datetime(queries['start_date'], unit='D', origin='1899-12-30', errors='coerce')

    # Refresh Button (User-Friendly)
    if st.button("🔄 Refresh Filters & Data"):
        st.rerun()

    # Hero Image (Creative UI)
    st.image("https://via.placeholder.com/1200x250/FFEFEF/FF5A6E?text=✨+Lady+Care+Glow-Up+Insights+(Sep+17,+2025)", use_column_width=True)

    colA, colB = st.columns([2,1])
    with colA:
        # Trend Line (Accuracy: Handle empty)
        counts_trend = queries.groupby('Date')['Counts'].sum().reset_index()
        if not counts_trend.empty and len(counts_trend) > 1:
            fig = px.line(counts_trend, x='Date', y='Counts', title='Counts over Time (2025 Trends)', 
                          labels={'Counts':'Search Counts'}, color_discrete_sequence=px.colors.qualitative.D3)
            fig.update_xaxes(title="Date (Serial to Datetime Fixed)")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("📅 Add more date range for trends. Sample: Q4 2025 shows INTIMATE CARE spike.")
    with colB:
        st.markdown("**Top 10 Queries (Counts)**")
        top10 = queries.nlargest(10, 'Counts')[['search', 'Counts', 'clicks', 'Converion Rate']].rename(columns={'search':'Query', 'Counts':'Search Counts'})
        st.dataframe(top10, use_container_width=True, help="From queries_clustered: High-vol like 'feminine wash' (1,642 Counts).")

    st.markdown("---")
    st.subheader("📊 Performance Snapshot")

    # Enhanced KPIs with st.metric (Animated, Tooltips)
    total_counts = int(queries['Counts'].sum()) if not queries['Counts'].empty else 0
    total_clicks = int(queries['clicks'].sum()) if not queries['clicks'].empty else 0
    total_conv_safe = int((queries['clicks'] * queries['Converion Rate']).sum().fillna(0))
    overall_ctr = (total_clicks / total_counts * 100) if total_counts > 0 else 0
    overall_cr = (total_conv_safe / total_clicks * 100) if total_clicks > 0 else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("✨ Total Counts", f"{total_counts:,}", help=f"From 'total_impressions over 3m' col (~{total_counts:,} parsed; full: 17M+ per summaries).")
    with col2:
        st.metric("👆 Total Clicks", f"{total_clicks:,}", help="From 'count' col (e.g., 3,802 partial; accurate daily clicks).")
    with col3:
        st.metric("🎯 Total Conversions", f"{total_conv_safe:,}", help="Clicks * CR (handles NaN; partial ~684).")
    with col4:
        st.metric("📈 Overall CTR", f"{overall_ctr:.2f}%", help="Clicks / Counts (e.g., 25.9% partial; category highs in FAMILY PLANNING @31.7%).")
    with col5:
        st.metric("💡 Overall CR", f"{overall_cr:.2f}%", help="Conversions / Clicks (e.g., 18.0% partial; classical_cr avg 1.0x base).")

    # Mini-Metrics Row (Data-Driven: From Analysis with Share)
    colM1, colM2, colM3, colM4 = st.columns(4)
    with colM1:
        avg_ctr = queries['Click Through Rate'].mean() * 100 if not queries.empty else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>📊</span>
            <div class='value'>{avg_ctr:.2f}%</div>
            <div class='label'>Avg CTR (All Cats)</div>
        </div>
        """, unsafe_allow_html=True)
    with colM2:
        avg_cr = queries['Converion Rate'].mean() * 100 if not queries.empty else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>🎯</span>
            <div class='value'>{avg_cr:.2f}%</div>
            <div class='label'>Avg CR (Derived)</div>
        </div>
        """, unsafe_allow_html=True)
    with colM3:
        unique_queries = queries['search'].nunique()
        total_share = (queries.groupby('search')['Counts'].sum() / total_counts * 100).max() if total_counts > 0 else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>🔍</span>
            <div class='value'>{unique_queries:,} ({total_share:.2f}%)</div>
            <div class='label'>Unique Queries (Top Share)</div>
        </div>
        """, unsafe_allow_html=True)
    with colM4:
        cat_counts = queries.groupby('Category')['Counts'].sum()
        top_cat = cat_counts.idxmax()
        top_cat_share = (cat_counts.max() / total_counts * 100) if total_counts > 0 else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>📦</span>
            <div class='value'>{int(cat_counts.max()):,} ({top_cat_share:.2f}%)</div>
            <div class='label'>Top Cat Counts ({top_cat})</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("🏷 Brand & Category Snapshot")
    g1, g2 = st.columns(2)
    with g1:
        if 'Brand' in queries.columns:
            brand_perf = queries.groupby('Brand').agg({'Counts':'sum', 'clicks':'sum', 'Converion Rate':'mean'}).reset_index()
            brand_perf['conversions'] = (brand_perf['clicks'] * brand_perf['Converion Rate']).round()
            brand_perf['share'] = (brand_perf['Counts'] / total_counts * 100).round(2)
            fig = px.bar(brand_perf.sort_values('Counts', ascending=False).head(10), 
                         x='Brand', y='Counts', title='Top Brands by Counts (e.g., Sofy Leads @614k Full)',
                         color='conversions', color_continuous_scale='Viridis', hover_data=['share'])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("🏷 Brand data ready from sheet.")
    with g2:
        if 'Category' in queries.columns:
            cat_perf = queries.groupby('Category').agg({'Counts':'sum', 'clicks':'sum', 'Converion Rate':'mean'}).reset_index()
            cat_perf['conversions'] = (cat_perf['clicks'] * cat_perf['Converion Rate']).round()
            cat_perf['share'] = (cat_perf['Counts'] / total_counts * 100).round(2)
            st.markdown("**Top Categories by Counts**")
            if AGGRID_OK:
                AgGrid(cat_perf.sort_values('Counts', ascending=False).head(10), height=300, enable_enterprise_modules=False)
            else:
                st.dataframe(cat_perf, use_container_width=True)
        else:
            st.info("📦 Category data parsed (e.g., SANITARY CARE: 6M+ Counts).")

# ----------------- Search Analysis (core) -----------------
with tab_search:
    st.header("🔍 Search Column — Deep Dive")
    st.markdown("Analyze raw search queries (no spelling corrections) with keyword frequency, query length, and long-tail insights. 🎯")

    kw_series = queries['keywords'].explode().dropna()
    kw_counts = kw_series.value_counts().reset_index()
    kw_counts.columns = ['keyword', 'count']  # Explicitly rename columns
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
    ql = queries.groupby('query_length').agg(Counts=('Counts','sum'), clicks=('clicks','sum')).reset_index()
    ql['ctr'] = ql.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
    if not ql.empty:
        st.plotly_chart(px.scatter(ql, x='query_length', y='ctr', size='Counts', title='Query Length vs CTR (Size=Counts)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
    else:
        st.info("No query length data.")

    st.subheader("📊 Long-Tail vs Short-Tail")
    queries['is_long_tail'] = queries['query_length'] >= 20
    lt = queries.groupby('is_long_tail').agg(Counts=('Counts','sum'), conversions=('conversions','sum')).reset_index()
    lt['label'] = lt['is_long_tail'].map({True:'Long-tail (>=20 chars)', False:'Short-tail'})
    if not lt.empty:
        st.plotly_chart(px.pie(lt, names='label', values='Counts', title='Counts Share: Long-Tail vs Short-Tail', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
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
        bs = queries.groupby('brand').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        bs['ctr'] = bs.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        bs['cr'] = bs.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        st.plotly_chart(px.bar(bs.sort_values('Counts', ascending=False).head(20), x='brand', y='Counts', title='Top Brands by Counts', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        st.plotly_chart(px.scatter(bs, x='Counts', y='ctr', size='conversions', color='brand', title='Brand: Counts vs CTR (Size=Conversions)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)

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
        cs = queries.groupby('category').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        cs['ctr'] = cs.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        cs['cr'] = cs.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        st.plotly_chart(px.bar(cs.sort_values('Counts', ascending=False), x='category', y='Counts', title='Counts by Category', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
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
        sc = queries.groupby('sub_category').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        sc['ctr'] = sc.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        st.plotly_chart(px.bar(sc.sort_values('Counts', ascending=False).head(30), x='sub_category', y='Counts', title='Top Subcategories by Counts', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
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

# ----------------- Time Analysis Tab (Modified) -----------------
with tab_time:
    st.header("⏰ Temporal Analysis & Seasonality")
    st.markdown("Uncover monthly trends to optimize campaigns. 📅")

    if queries['month'].notna().any():
        monthly = queries.groupby('month').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        monthly['ctr'] = monthly.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        try:
            monthly['month_dt'] = pd.to_datetime(monthly['month'], format='%b %Y', errors='coerce')
            monthly = monthly.sort_values('month_dt')
        except:
            monthly = monthly.sort_values('month')
        st.plotly_chart(px.line(monthly, x='month', y='Counts', title='Monthly Counts', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
        st.plotly_chart(px.line(monthly, x='month', y='ctr', title='Monthly Average CTR (%)', color_discrete_sequence=px.colors.qualitative.Plotly), use_container_width=True)
    else:
        st.info("No month data to plot.")

    st.subheader("🏷 Top Brands by Month (Counts)")  # Replaced CTR by Day of Week
    if 'brand' in queries.columns and queries['brand'].notna().any() and queries['month'].notna().any():
        top_brands = queries.groupby('brand')['Counts'].sum().sort_values(ascending=False).head(5).index
        brand_month = queries[queries['brand'].isin(top_brands)].groupby(['month','brand']).agg(Counts=('Counts','sum')).reset_index()
        try:
            brand_month['month_dt'] = pd.to_datetime(brand_month['month'], format='%b %Y', errors='coerce')
            brand_month = brand_month.sort_values('month_dt')
        except:
            brand_month = brand_month.sort_values('month')
        fig = px.bar(brand_month, x='month', y='Counts', color='brand', title='Top 5 Brands by Counts per Month', color_discrete_sequence=px.colors.qualitative.D3)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Brand or month data not available for brand-month analysis.")

# ----------------- Pivot Builder Tab -----------------
with tab_pivot:
    st.header("📊 Pivot Builder & Prebuilt Pivots")
    st.markdown("Create custom pivots or explore prebuilt tables for quick insights. 🔧")

    st.subheader("📋 Prebuilt: Brand × Query (Top 300)")
    if 'brand' in queries.columns:
        pv = queries.groupby(['brand','normalized_query']).agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        pv['ctr'] = pv.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        pv_top = pv.sort_values('Counts', ascending=False).head(300)
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
    val = st.selectbox("Value (Measure)", options=['Counts','clicks','conversions'], index=0)  # Updated to Counts
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

# ----------------- Insights & Questions (Modified) -----------------
with tab_insights:
    st.header("💡 Insights & Actionable Questions (26)")
    st.markdown("Actionable insights focused on the **search** column, with pivot tables and visuals. 🚀")

    def q_expand(title, explanation, render_fn, icon="💡"):
        with st.expander(f"{icon} {title}", expanded=False):
            st.markdown(f"<div class='insight-box'><h4>Why & How to Use</h4><p>{explanation}</p></div>", unsafe_allow_html=True)
            try:
                render_fn()
            except Exception as e:
                st.error(f"Rendering error: {e}")

    # Q1: Top queries by Counts (originally Q1)
    def q1():
        out = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('Counts', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q1 — Top Queries by Counts (Top 30)", "Which queries drive the most traffic? Prioritize for search tuning and inventory.", q1, "📈")

    # Q2: High Counts, low CTR (originally Q2)
    def q2():
        df2 = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum')).reset_index()
        df2['ctr'] = df2.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        out = df2[(df2['Counts']>=df2['Counts'].quantile(0.6)) & (df2['ctr']<=df2['ctr'].quantile(0.3))].sort_values('Counts', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q2 — High Counts, Low CTR", "Queries with high traffic but low engagement. Improve relevance, snippets, or imagery.", q2, "⚠️")

    # Q3: Top queries by conversion rate (originally Q4)
    def q3():
        df4 = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df4 = df4[df4['Counts']>=50]
        df4['cr'] = df4.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df4.sort_values('cr', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q3 — Top Queries by Conversion Rate (Min Counts=50)", "High-converting queries for paid promotions.", q3, "🎯")

    # Q4: Long-tail contribution (originally Q5)
    def q4():
        lt = queries[queries['query_length']>=20]
        st.markdown(f"Long-tail rows: {len(lt):,} / total {len(queries):,}")
        st.plotly_chart(px.pie(names=['Long-tail','Rest'], values=[lt['Counts'].sum(), queries['Counts'].sum()-lt['Counts'].sum()], title='Counts Share: Long-Tail vs Rest', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    q_expand("Q4 — Long-Tail vs Short-Tail (>=20 chars)", "How much traffic comes from long-tail queries? Key for content strategy.", q4, "📏")

    # Q5: Brand vs generic share (originally Q7)
    def q5():
        if 'brand' in queries.columns:
            branded = queries[queries['brand'].notna() & (queries['brand']!='')]
            branded_share = branded['Counts'].sum()
            total = queries['Counts'].sum()
            st.markdown(f"Branded Counts: {branded_share:,} / Total: {total:,}  —  Share: {branded_share/total:.2%}")
            st.plotly_chart(px.pie(names=['Branded','Generic'], values=[branded_share, total-branded_share], title='Branded vs Generic Counts Share', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
            pivot = queries.pivot_table(values=['Counts','clicks','conversions'], index=['brand'], aggfunc='sum').reset_index()
            if AGGRID_OK:
                AgGrid(pivot, height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Brand column not present.")
    q_expand("Q5 — Branded vs Generic Queries (Pivot)", "Assess brand vs generic search intent with a pivot table.", q5, "🏷")

    # Q6: Rising queries MoM (originally Q8)
    def q6():
        mom = queries.groupby(['month','normalized_query']).agg(Counts=('Counts','sum')).reset_index()
        if len(mom['month'].unique())<2:
            st.info("Not enough months to compute MoM.")
            return
        pivot = mom.pivot(index='normalized_query', columns='month', values='Counts').fillna(0)
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
    q_expand("Q6 — Top Rising Queries Month-over-Month", "Detect emerging demand for seasonal campaigns.", q6, "📈")

    # Q7: Query funnel snapshot (originally Q11)
    def q7():
        snap = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('Counts', ascending=False).head(200)
        if AGGRID_OK:
            AgGrid(snap, height=400)
        else:
            st.dataframe(snap.head(100), use_container_width=True)
    q_expand("Q7 — Query Funnel Snapshot (Top 200)", "View top queries' funnel: Counts → clicks → conversions.", q7, "📊")

    # Q8: Traffic concentration (originally Q14)
    def q8():
        qq = queries.groupby('normalized_query').agg(Counts=('Counts','sum')).reset_index().sort_values('Counts', ascending=False)
        top5n = max(1, int(0.05*len(qq)))
        share = qq.head(top5n)['Counts'].sum() / qq['Counts'].sum() if qq['Counts'].sum()>0 else 0
        st.markdown(f"Top 5% queries contribute **{share:.2%}** of Counts (top {top5n} queries).")
        st.plotly_chart(px.pie(names=['Top 5% Queries','Rest'], values=[qq.head(top5n)['Counts'].sum(), qq['Counts'].sum()-qq.head(top5n)['Counts'].sum()], title='Traffic Concentration: Top 5% Queries', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    q_expand("Q8 — Traffic Concentration (Top 5%)", "Prioritize top queries driving most traffic.", q8, "📈")

    # Q9: Keyword co-occurrence (originally Q15)
    def q9():
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
    q_expand("Q9 — Keyword Co-Occurrence (Cross-Sell Proxy)", "Find keywords searched together for cross-sell opportunities.", q9, "🔗")

    # Q10: High searches, zero conversions (originally Q16)
    def q10():
        dfm = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        out = dfm[(dfm['Counts']>=dfm['Counts'].quantile(0.7)) & (dfm['conversions']==0)].sort_values('Counts', ascending=False).head(40)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q10 — High Search Volume, Zero Conversions", "Fix product discovery or pricing for these queries.", q10, "⚠️")

    # Q11: Queries with many variants (originally Q17)
    def q11():
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
    q_expand("Q11 — Queries with Many Variants (Prefix Clustering)", "Identify queries with variants/typos for canonicalization.", q11, "🔍")

    # Q12: Top queries by CTR (originally Q19)
    def q12():
        df19 = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum')).reset_index()
        df19 = df19[df19['Counts']>=30]
        df19['ctr'] = df19.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        out = df19.sort_values('ctr', ascending=False).head(40)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q12 — Top Queries by CTR (Min Counts=30)", "High-engagement queries for ad campaigns.", q12, "📈")

    # Q13: Low CTR & CR queries (originally Q20)
    def q13():
        df20 = queries.groupby('normalized_query').agg(Counts=('Counts','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df20['ctr'] = df20.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        df20['cr'] = df20.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df20[(df20['Counts']>=df20['Counts'].quantile(0.6)) & (df20['ctr']<=df20['ctr'].quantile(0.25)) & (df20['cr']<=df20['cr'].quantile(0.25))].sort_values('Counts', ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q13 — High Counts, Low CTR & CR", "Optimize search results for these underperforming queries.", q13, "⚠️")

    # Q14: Top keywords per category (originally Q21)
    def q14():
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
    q_expand("Q14 — Top Keywords per Category (Pivot)", "Understand category-specific search language for taxonomy.", q14, "📦")

    # Q15: Brand-inclusive queries (originally Q23)
    def q15():
        if 'brand' in queries.columns:
            labeled = queries[queries['brand'].notna() & (queries['brand']!='')]
            brand_q = labeled.groupby('normalized_query').size().reset_index(name='count').sort_values('count', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(brand_q, height=400)
            else:
                st.dataframe(brand_q, use_container_width=True)
        else:
            st.info("Brand column missing.")
    q_expand("Q15 — Brand-Inclusive Queries", "High purchase intent queries with brands.", q15, "🏷")

    # Q16: Top queries by conversions (originally Q26)
    def q16():
        out = queries.groupby('normalized_query').agg(conversions=('conversions','sum'), Counts=('Counts','sum')).reset_index().sort_values('conversions', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q16 — Top Queries by Conversions", "Direct revenue drivers for inventory and bids.", q16, "🎯")

    # Q17: Month-by-month query trends (originally Q27)
    def q17():
        mom = queries.groupby(['month','normalized_query']).agg(Counts=('Counts','sum')).reset_index()
        topq = mom.groupby('normalized_query')['Counts'].sum().reset_index().sort_values('Counts', ascending=False).head(500)['normalized_query']
        sample = mom[mom['normalized_query'].isin(topq)].pivot(index='normalized_query', columns='month', values='Counts').fillna(0)
        if sample.shape[1] >= 2:
            if AGGRID_OK:
                AgGrid(sample.head(200).reset_index(), height=400)
            else:
                st.dataframe(sample.head(200), use_container_width=True)
        else:
            st.info("Not enough months to show seasonality.")
    q_expand("Q17 — Month-by-Month Query Trends (Pivot)", "Identify seasonal trends for campaign planning.", q17, "📅")

    # Q18: Unique keywords by category (originally Q29)
    def q18():
        if 'category' in queries.columns:
            uniq = queries.groupby('category').agg(unique_keywords=('keywords', lambda s: len(set([k for sub in s for k in sub])))).reset_index().sort_values('unique_keywords', ascending=False)
            if AGGRID_OK:
                AgGrid(uniq, height=400)
            else:
                st.dataframe(uniq, use_container_width=True)
        else:
            st.info("Category missing.")
    q_expand("Q18 — Unique Keywords by Category", "Measure search diversity for facet planning.", q18, "📦")

    # Q19: Long-term growth queries (originally Q30)
    def q19():
        if queries['month'].nunique() < 2:
            st.info("Not enough months.")
            return
        months_sorted = sorted(queries['month'].dropna().unique(), key=lambda x: pd.to_datetime(x, format='%b %Y', errors='coerce'))
        first, last = months_sorted[0], months_sorted[-1]
        m1 = queries[queries['month']==first].groupby('normalized_query').agg(Counts=('Counts','sum')).rename(columns={'Counts':'m1'})
        m2 = queries[queries['month']==last].groupby('normalized_query').agg(Counts=('Counts','sum')).rename(columns={'Counts':'m2'})
        comp = m1.join(m2, how='outer').fillna(0)
        comp['pct_change'] = (comp['m2'] - comp['m1']) / comp['m1'].replace(0, np.nan)
        out = comp.sort_values('pct_change', ascending=False).head(50).reset_index()
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q19 — Long-Term Growth: First vs Last Month", "Find queries with significant growth or decline.", q19, "📈")

    # Q20: Top 50 queries (quick) (originally Q31)
    def q20():
        out = queries.groupby('normalized_query').agg(Counts=('Counts','sum')).reset_index().sort_values('Counts',ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q20 — Top 50 Queries (Quick)", "Quick ranking of top queries by Counts.", q20, "📊")

    # Q21: Top brands quick view (originally Q32)
    def q21():
        if 'brand' in queries.columns:
            out = queries.groupby('brand').agg(Counts=('Counts','sum'), conversions=('conversions','sum')).reset_index().sort_values('Counts', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Brand missing.")
    q_expand("Q21 — Top Brands (Quick)", "Quick brand ranking by Counts.", q21, "🏷")

    # Q22: Top subcategories quick view (originally Q33)
    def q22():
        if 'sub_category' in queries.columns:
            out = queries.groupby('sub_category').agg(Counts=('Counts','sum')).reset_index().sort_values('Counts', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Subcategory missing.")
    q_expand("Q22 — Top Subcategories (Quick)", "Quick subcategory ranking.", q22, "🧴")

    # Q23: Monthly Counts table (originally Q35)
    def q23():
        out = queries.groupby('month').agg(Counts=('Counts','sum')).reset_index().sort_values('month')
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q23 — Monthly Counts Table", "Month-level volumes for reporting.", q23, "📅")

    # Q24: Top queries by clicks (originally Q37)
    def q24():
        out = queries.groupby('normalized_query').agg(clicks=('clicks','sum'), Counts=('Counts','sum')).reset_index().sort_values('clicks', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q24 — Top Queries by Clicks", "High-engagement queries for ad optimization.", q24, "👆")

    # Q25: Category vs brand performance (originally Q38)
    def q25():
        if 'category' in queries.columns and 'brand' in queries.columns:
            pivot = queries.pivot_table(values=['Counts','clicks','conversions'], index=['category'], columns=['brand'], aggfunc='sum').fillna(0)
            if AGGRID_OK:
                AgGrid(pivot.reset_index(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Category or brand missing.")
    q_expand("Q25 — Category vs Brand Performance (Pivot)", "Analyze brand performance within categories.", q25, "📦🏷")

    # Q26: High Counts, low clicks by category (originally Q39)
    def q26():
        if 'category' in queries.columns:
            df39 = queries.groupby(['category','normalized_query']).agg(Counts=('Counts','sum'), clicks=('clicks','sum')).reset_index()
            df39['ctr'] = df39.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
            out = df39[(df39['Counts']>=df39['Counts'].quantile(0.6)) & (df39['ctr']<=df39['ctr'].quantile(0.3))].sort_values('Counts', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Category missing.")
    q_expand("Q26 — High Counts, Low Clicks by Category", "Identify category-specific queries needing optimization.", q26, "⚠️")

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
✨ Lady Care Search Analytics — Noureldeen Mohamed
</div>
""", unsafe_allow_html=True)