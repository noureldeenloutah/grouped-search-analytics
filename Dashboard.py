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

/* Mini metric cards */
.mini-metric {
    background: #FFFFFF;
    padding: 12px;
    border-radius: 8px;
    text-align: center;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    margin: 5px;
}
.mini-metric .value {
    font-size: 1.2rem;
    font-weight: 700;
    color: #FF5A6E;
}
.mini-metric .label {
    font-size: 0.8rem;
    color: #6C7A89;
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
            df['Date'] = pd.to_datetime(df['start_date'], errors='coerce')
    else:
        df['Date'] = pd.NaT

    # Impressions
    if 'total_impressions over 3m' in df.columns:
        df['impressions'] = pd.to_numeric(df['total_impressions over 3m'], errors='coerce').fillna(0)
    else:
        df['impressions'] = 0

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
        df['ctr'] = df.apply(lambda r: r['clicks']/r['impressions'] if r['impressions']>0 else 0, axis=1) * 100
    
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
        return df[df[col].astize(str).isin(sel)]

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

total_impr = int(queries['impressions'].sum())
total_clicks = int(queries['clicks'].sum())
total_conv = int(queries['conversions'].sum())
overall_ctr = (queries['clicks'].sum()/queries['impressions'].sum()) * 100 if queries['impressions'].sum()>0 else 0
overall_cr = (queries['conversions'].sum()/queries['clicks'].sum()) * 100 if queries['clicks'].sum()>0 else 0
total_revenue = 0.0  # No revenue column

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
    st.subheader("📊 Performance Snapshot")
    
    # Mini metrics row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        avg_ctr = queries['ctr'].mean() if 'ctr' in queries.columns else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <div class='value'>{avg_ctr:.2f}%</div>
            <div class='label'>📊 Avg CTR</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        avg_cr = queries['cr'].mean() if 'cr' in queries.columns else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <div class='value'>{avg_cr:.2f}%</div>
            <div class='label'>🎯 Avg CR</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        unique_queries = queries['normalized_query'].nunique()
        st.markdown(f"""
        <div class='mini-metric'>
            <div class='value'>{unique_queries:,}</div>
            <div class='label'>🔍 Unique Queries</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        avg_query_len = queries['query_length'].mean() if 'query_length' in queries.columns else 0
        st.markdown(f"""
        <div class='mini-metric'>
            <div class='value'>{avg_query_len:.1f}</div>
            <div class='label'>📏 Avg Query Length</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Performance by category
    if 'category' in queries.columns and queries['category'].notna().any():
        st.subheader("📦 Performance by Category")
        cat_perf = queries.groupby('category').agg(
            impressions=('impressions', 'sum'),
            clicks=('clicks', 'sum'),
            conversions=('conversions', 'sum')
        ).reset_index()
        cat_perf['ctr'] = (cat_perf['clicks'] / cat_perf['impressions'] * 100).round(2)
        cat_perf['cr'] = (cat_perf['conversions'] / cat_perf['clicks'] * 100).round(2)
        
        fig = px.bar(cat_perf.sort_values('impressions', ascending=False).head(10), 
                    x='category', y='impressions', title='Top Categories by Impressions',
                    color='ctr', color_continuous_scale='Viridis')
        st.plotly_chart(fig, use_container_width=True)

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
    st.markdown("Uncover monthly trends to optimize campaigns. 📅")

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
        
        # Brand vs Month performance
        if 'brand' in queries.columns and queries['brand'].notna().any():
            st.subheader("🏷 Brand Performance by Month")
            brand_month = queries.groupby(['month', 'brand']).agg(impressions=('impressions','sum')).reset_index()
            top_brands = brand_month.groupby('brand')['impressions'].sum().nlargest(5).index.tolist()
            brand_month_top = brand_month[brand_month['brand'].isin(top_brands)]
            
            fig = px.line(brand_month_top, x='month', y='impressions', color='brand', 
                         title='Top 5 Brands: Monthly Impressions', color_discrete_sequence=px.colors.qualitative.D3)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No month data to plot.")

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
    idx = st.mult