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

import pandas as pd

import pandas as pd
import streamlit as st

def prepare_queries_df(df: pd.DataFrame, use_derived_metrics: bool = False):
    """Normalize columns, create derived metrics and time buckets.
    
    Args:
        df (pd.DataFrame): Input DataFrame from Excel sheet.
        use_derived_metrics (bool): If True, derive clicks and conversions from rates; if False, use sheet columns.
    """
    df = df.copy()
    
    # -------------------------
    # Query text
    # -------------------------
    if 'search' in df.columns:
        df['normalized_query'] = df['search'].astype(str)
    else:
        df['normalized_query'] = df.iloc[:, 0].astype(str)

    # -------------------------
    # Date normalization
    # -------------------------
    if 'start_date' in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df['start_date']):
            df['Date'] = df['start_date']
        else:
            df['Date'] = pd.to_datetime(
                df['start_date'], unit='D', origin='1899-12-30', errors='coerce'
            )
    else:
        df['Date'] = pd.NaT

    # -------------------------
    # COUNTS = search counts (from 'count' column)
    # -------------------------
    if 'count' in df.columns:
        df['Counts'] = pd.to_numeric(df['count'], errors='coerce').fillna(0)
        st.sidebar.success(f"✅ Using 'count' column for impressions: {df['Counts'].sum():,}")
    else:
        df['Counts'] = 0
        st.sidebar.warning("❌ No 'count' column found for impressions")

    # -------------------------
    # CLICKS and CONVERSIONS (use sheet columns or derive from rates)
    # -------------------------
    if 'Clicks' in df.columns:
        df['clicks'] = pd.to_numeric(df['Clicks'], errors='coerce').fillna(0)
        st.sidebar.success(f"✅ Using 'Clicks' column: {df['clicks'].sum():,}")
    else:
        df['clicks'] = 0
        st.sidebar.warning("❌ No 'Clicks' column found")

    if 'Conversions' in df.columns:
        df['conversions'] = pd.to_numeric(df['Conversions'], errors='coerce').fillna(0)
        st.sidebar.success(f"✅ Using 'Conversions' column: {df['conversions'].sum():,}")
    else:
        df['conversions'] = 0
        st.sidebar.warning("❌ No 'Conversions' column found")

    # Derive metrics if requested (overrides sheet values)
    if use_derived_metrics:
        if 'Click Through Rate' in df.columns and 'count' in df.columns:
            ctr = pd.to_numeric(df['Click Through Rate'], errors='coerce').fillna(0)
            if ctr.max() > 1:  # Percentage format
                ctr_decimal = ctr / 100.0
            else:  # Decimal format
                ctr_decimal = ctr
            df['clicks'] = (df['Counts'] * ctr_decimal).round().astype(int)
            st.sidebar.success(f"✅ Derived clicks from CTR: {df['clicks'].sum():,}")
        else:
            st.sidebar.warning("❌ Cannot derive clicks - missing CTR or count data")

        if 'Conversion Rate' in df.columns:  # Fixed typo from 'Converion Rate'
            conv_rate = pd.to_numeric(df['Conversion Rate'], errors='coerce').fillna(0)
            if conv_rate.max() > 1:  # Percentage format
                conv_rate_decimal = conv_rate / 100.0
            else:  # Decimal format
                conv_rate_decimal = conv_rate
            df['conversions'] = (df['clicks'] * conv_rate_decimal).round().astype(int)
            st.sidebar.success(f"✅ Derived conversions: {df['conversions'].sum():,}")
        else:
            st.sidebar.warning("❌ No Conversion Rate data found")

    # Validate derived vs. sheet values (if both exist)
    if 'Clicks' in df.columns and use_derived_metrics:
        diff_clicks = abs(df['clicks'].sum() - df['Clicks'].sum())
        if diff_clicks > 0:
            st.sidebar.warning(f"⚠ Derived clicks ({df['clicks'].sum():,}) differ from sheet Clicks ({df['Clicks'].sum():,}) by {diff_clicks:,}")
    if 'Conversions' in df.columns and use_derived_metrics:
        diff_conversions = abs(df['conversions'].sum() - df['Conversions'].sum())
        if diff_conversions > 0:
            st.sidebar.warning(f"⚠ Derived conversions ({df['conversions'].sum():,}) differ from sheet Conversions ({df['Conversions'].sum():,}) by {diff_conversions:,}")

    # -------------------------
    # CTR (store as percentage for consistency)
    # -------------------------
    if 'Click Through Rate' in df.columns:
        ctr = pd.to_numeric(df['Click Through Rate'], errors='coerce').fillna(0)
        if ctr.max() <= 1:
            df['ctr'] = ctr * 100  # Convert to percentage
        else:
            df['ctr'] = ctr  # Already in percentage
    else:
        df['ctr'] = df.apply(
            lambda r: (r['clicks'] / r['Counts']) * 100 if r['Counts'] > 0 else 0, axis=1
        )

    # -------------------------
    # CR (store as percentage for consistency)
    # -------------------------
    if 'Conversion Rate' in df.columns:  # Fixed typo
        cr = pd.to_numeric(df['Conversion Rate'], errors='coerce').fillna(0)
        if cr.max() <= 1:
            df['cr'] = cr * 100  # Convert to percentage
        else:
            df['cr'] = cr  # Already in percentage
    else:
        df['cr'] = df.apply(
            lambda r: (r['conversions'] / r['Counts']) * 100 if r['Counts'] > 0 else 0,
            axis=1,
        )

    # Classical CR
    if 'classical_cr' in df.columns:
        classical_cr = pd.to_numeric(df['classical_cr'], errors='coerce').fillna(0)
        if classical_cr.max() <= 1:
            df['classical_cr'] = classical_cr * 100
        else:
            df['classical_cr'] = classical_cr
    else:
        df['classical_cr'] = df['cr']

    # -------------------------
    # Revenue (placeholder)
    # -------------------------
    df['revenue'] = 0

    # -------------------------
    # Time buckets
    # -------------------------
    df['year'] = df['Date'].dt.year
    df['month'] = df['Date'].dt.strftime('%B %Y')
    df['month_short'] = df['Date'].dt.strftime('%b')
    df['day_of_week'] = df['Date'].dt.day_name()

    # -------------------------
    # Text features
    # -------------------------
    df['query_length'] = df['normalized_query'].astype(str).apply(len)
    df['keywords'] = df['normalized_query'].apply(extract_keywords)  # Assuming extract_keywords is defined

    # -------------------------
    # Brand, Category, Subcategory, Department
    # -------------------------
    df['brand_ar'] = ''
    df['brand'] = df['Brand'] if 'Brand' in df.columns else None
    df['category'] = df['Category'] if 'Category' in df.columns else None
    df['sub_category'] = df['Sub Category'] if 'Sub Category' in df.columns else None
    df['department'] = df['Department'] if 'Department' in df.columns else None

    # -------------------------
    # Additional optional columns
    # -------------------------
    if 'underperforming' in df.columns:
        df['underperforming'] = df['underperforming']
    if 'averageClickPosition' in df.columns:
        df['average_click_position'] = df['averageClickPosition']
    if 'cluster_id' in df.columns:
        df['cluster_id'] = df['cluster_id']

    # -------------------------
    # Keep original columns for reference
    # -------------------------
    original_cols = ['Department', 'Category', 'Sub Category', 'Brand', 'search', 'count', 
                     'Click Through Rate', 'Conversion Rate', 'total_impressions over 3m',
                     'averageClickPosition', 'underperforming', 'classical_cr', 'cluster_id',
                     'start_date', 'end_date']
    
    for col in original_cols:
        if col in df.columns:
            df[f'orig_{col}'] = df[col]

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

# Calculate metrics
total_counts = int(queries['Counts'].sum())
total_clicks = int(queries['clicks'].sum())
total_conversions = int(queries['conversions'].sum())
overall_ctr = (total_clicks / total_counts * 100) if total_counts > 0 else 0
overall_cr = (total_conversions / total_counts * 100) if total_clicks > 0 else 0
total_revenue = 0.0  # No revenue column

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f"<div class='kpi'><div class='value'>{total_counts:,}</div><div class='label'>✨ Total Counts</div></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='kpi'><div class='value'>{total_clicks:,}</div><div class='label'>👆 Total Clicks</div></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='kpi'><div class='value'>{total_conversions:,}</div><div class='label'>🎯 Total Conversions</div></div>", unsafe_allow_html=True)
with c4:
    st.markdown(f"<div class='kpi'><div class='value'>{overall_ctr:.2f}%</div><div class='label'>📈 Overall CTR</div></div>", unsafe_allow_html=True)
with c5:
    st.markdown(f"<div class='kpi'><div class='value'>{overall_cr:.2f}%</div><div class='label'>💡 Overall CR</div></div>", unsafe_allow_html=True)

# Show data source info in sidebar
st.sidebar.info(f"**Data Source:** {main_key}")
st.sidebar.write(f"**Total Rows:** {len(queries):,}")
st.sidebar.write(f"**Total Counts:** {total_counts:,}")
st.sidebar.write(f"**Calculated Clicks:** {total_clicks:,}")
st.sidebar.write(f"**Calculated Conversions:** {total_conversions:,}")

# Add debug info in an expander so it doesn't clutter the sidebar
with st.sidebar.expander("🔍 Data Debug Info"):
    st.write(f"Main sheet: {main_key}")
    st.write(f"Processed columns: {list(queries.columns)}")
    st.write(f"Processed shape: {queries.shape}")
    
    st.write("**Column Usage:**")
    if 'count' in raw_queries.columns:
        st.write(f"✓ Counts/Impressions: 'count' column")
    else:
        st.write("✗ Counts/Impressions: No 'count' column found")
    
    st.write("**Calculation Method:**")
    st.write("• Clicks = Counts × Click Through Rate")
    st.write("• Conversions = Clicks × Conversion Rate")
    
    # Show sample of raw data
    st.write("**Sample data (first 3 rows):**")
    st.dataframe(raw_queries.head(3))

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

    # Image Selection in Sidebar
    st.sidebar.header("🎨 Customize Hero Image")
    image_options = {
        "Abstract Gradient": "https://placehold.co/1200x250/FF5A6E/FFFFFF?text=✨+Lady+Care+Insights",
        "Nature-Inspired": "https://picsum.photos/1200/250?random=care_nature",
        "Elegant Pink Theme": "https://source.unsplash.com/1200x250/?pink,elegant",
        "Custom Text on Solid Color": "https://placehold.co/1200x250/E6F3FA/FF5A6E?text=✨+Lady+Care+Glow",
        "Feminine Floral": "https://picsum.photos/1200/250?random=floral_feminine"
    }
    selected_image = st.sidebar.selectbox("Choose Hero Image", options=list(image_options.keys()), index=0)

    # Hero Image (Creative UI) with selected option
    st.image(image_options[selected_image], use_container_width=True)

    colA, colB = st.columns([2, 1])
    with colA:
        # Counts over Months as a creative bar chart with labels and percentages
        monthly_counts = queries.groupby(queries['Date'].dt.strftime('%B %Y'))['Counts'].sum().reset_index()
        st.write("Monthly Counts:", monthly_counts)  # Debug to verify sums
        if not monthly_counts.empty and len(monthly_counts) >= 2:
            # Ensure 'Counts' is numeric and handle NaN
            monthly_counts['Counts'] = pd.to_numeric(monthly_counts['Counts'], errors='coerce').fillna(0)
            total_all_months = monthly_counts['Counts'].sum()
            monthly_counts['Percentage'] = (monthly_counts['Counts'] / total_all_months * 100).round(1)

            # Create a beautiful bar chart with text labels
            try:
                fig = px.bar(monthly_counts, x='Date', y='Counts',
                            title='<b style="color:#FF5A6E; font-size:18px; text-shadow: 2px 2px 4px #00000055;">Counts Over Months: 2025 Trends at a Glance! 🌟</b>',
                            labels={'Date': '<i>Month</i>', 'Counts': '<b>Search Counts</b>'},
                            color='Counts',
                            color_continuous_scale=['#E6F3FA', '#FFB085', '#FF5A6E'],
                            template='plotly_white',
                            text=monthly_counts['Counts'].astype(str))  # Show counts on bars
                    
                # Update traces to position text outside, add percentages, and set hovertemplate
                fig.update_traces(
                    texttemplate='%{text}<br>%{customdata:.1f}%',
                    customdata=monthly_counts['Percentage'],
                    textposition='outside',
                    hovertemplate='<b>%{x}</b><br>Counts: %{y:,.0f}<br>Share: %{customdata:.1f}%<extra></extra>'
                )
                
                # Enhance attractiveness: Custom layout for beauty
                fig.update_layout(
                    plot_bgcolor='rgba(255,255,255,0.95)',
                    paper_bgcolor='rgba(255,247,232,0.8)',
                    font=dict(color='#0B486B', family='Segoe UI'),
                    title_x=0,  # Left alignment for title
                    title_font_size=16,
                    xaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
                    yaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
                    bargap=0.2,
                    barcornerradius=8,
                    annotations=[
                        dict(
                            x=0.5, y=1.05, xref='paper', yref='paper',
                            text='✨ Hover for details | Peak month highlighted below ✨',
                            showarrow=False,
                            font=dict(size=10, color='#FF5A6E', family='Segoe UI'),
                            align='center'
                        )
                    ]
                )
                
                # Highlight the peak month with a custom marker or annotation
                peak_month = monthly_counts.loc[monthly_counts['Counts'].idxmax(), 'Date']
                peak_value = monthly_counts['Counts'].max()
                fig.add_annotation(
                    x=peak_month, y=peak_value,
                    text=f"🏆 Peak: {peak_value:,.0f}",
                    showarrow=True,
                    arrowhead=3,
                    arrowcolor='#FF5A6E',
                    ax=0, ay=-30,
                    font=dict(size=12, color='#FF5A6E', family='Segoe UI', weight='bold')
                )
                
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generating chart: {e}")
        else:
            st.info("📅 Add more date range for monthly trends. Sample: Q4 2025 shows INTIMATE CARE spike.")

    with colB:
        st.markdown("**Top 50 Queries (Counts)**")
        if queries.empty or 'Counts' not in queries.columns or queries['Counts'].isna().all():
            st.warning("No valid data available for top 50 queries.")
        else:
            try:
                # Debug: Log available columns
                st.sidebar.write("**Available columns in queries:**", list(queries.columns))

                # Group by 'search' and aggregate
                top50 = queries.groupby('search').agg({
                    'Counts': 'sum',
                    'clicks': 'sum',
                    'conversions': 'sum'
                }).reset_index()

                # Calculate total Counts for share percentage
                total_counts = queries['Counts'].sum()

                # Calculate Conversion Rate based on conversions / Counts if column exists or as fallback
                if 'Conversion Rate' in queries.columns:
                    top50['Conversion Rate'] = pd.to_numeric(queries.groupby('search')['Conversion Rate'].mean(), errors='coerce').fillna(0)
                else:
                    # Derive Conversion Rate as (conversions / Counts * 100)
                    top50['Conversion Rate'] = (top50['conversions'] / top50['Counts'] * 100).round(2).fillna(0).replace([float('inf'), -float('inf')], 0)

                # Calculate share percentage
                top50['Share %'] = (top50['Counts'] / total_counts * 100).round(2)

                # Sort by 'Counts' and get top 50
                top50 = top50.nlargest(50, 'Counts')

                # Rename columns for display and format
                top50 = top50.rename(columns={
                    'search': 'Query',
                    'Counts': 'Search Counts',
                    'clicks': 'Clicks',
                    'conversions': 'Conversions'
                })

                # Round up clicks and format Conversion Rate as percentage
                top50['Clicks'] = top50['Clicks'].round().astype(int)
                top50['Conversion Rate'] = top50['Conversion Rate'].astype(str) + '%' if top50['Conversion Rate'].dtype != 'object' else top50['Conversion Rate']

                # Format Search Counts with commas
                top50['Search Counts'] = top50['Search Counts'].apply(lambda x: f"{x:,.0f}")

                # Reorder columns to place Share % after Search Counts
                column_order = ['Query', 'Search Counts', 'Share %', 'Clicks', 'Conversions', 'Conversion Rate']
                top50 = top50[column_order]

                # Center-align all values using Styler
                styled_top50 = top50.style.set_properties(**{
                    'text-align': 'center',
                    'font-size': '14px'
                }).format({
                    'Search Counts': '{}',
                    'Share %': '{:.2f}%',
                    'Clicks': '{:,.0f}',
                    'Conversions': '{:,.0f}',
                    'Conversion Rate': '{}'
                })

                # Display the DataFrame
                st.dataframe(styled_top50, use_container_width=True)

                # Add download button
                csv = top50.to_csv(index=False)
                st.download_button(
                    label="Download Table as CSV",
                    data=csv,
                    file_name="top_50_queries.csv",
                    mime="text/csv"
                )
            except KeyError as e:
                st.error(f"Column error: {e}. Check column names in your data (e.g., 'search', 'Counts', 'clicks', 'conversions', 'Conversion Rate').")
            except Exception as e:
                st.error(f"Error processing top 50 queries: {e}")

    st.markdown("---")
# ----------------- Performance Snapshot -----------------
    st.subheader("📊 Performance Snapshot")

    # Ensure columns exist before calculations
    required_cols_ps = ['Counts', 'clicks']
    if not all(col in queries.columns for col in required_cols_ps):
        st.error(f"Required columns {required_cols_ps} not found in dataset.")
        st.stop()

    # Calculate metrics with error handling (dynamic, updates with filters)
    total_counts = int(queries['Counts'].sum()) if not queries['Counts'].empty else 0
    total_clicks = int(queries['clicks'].sum()) if not queries['clicks'].empty else 0

    # Use 'conversions' column primarily for consistency with KPI Cards; fallback to Conversion Rate
    if 'conversions' in queries.columns:
        total_conv_safe = int(queries['conversions'].sum()) if not queries['conversions'].empty else 0
    elif 'Conversion Rate' in queries.columns:
        queries['Conversion Rate'] = pd.to_numeric(queries['Conversion Rate'], errors='coerce').fillna(0)
        total_conv_safe = int((queries['clicks'] * queries['Conversion Rate']).sum())
    else:
        total_conv_safe = 0  # Default to 0 if neither column exists

    overall_ctr = (total_clicks / total_counts * 100) if total_counts > 0 else 0
    overall_cr = (total_conv_safe / total_counts * 100) if total_clicks > 0 else 0

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown(f"<div class='kpi'><div class='value'>{total_counts:,}</div><div class='label'>✨ Total Counts</div></div>", unsafe_allow_html=True)

    with col2:
        st.markdown(f"<div class='kpi'><div class='value'>{total_clicks:,}</div><div class='label'>👆 Total Clicks</div></div>", unsafe_allow_html=True)

    with col3:
        st.markdown(f"<div class='kpi'><div class='value'>{total_conv_safe:,}</div><div class='label'>🎯 Total Conversions</div></div>", unsafe_allow_html=True)

    with col4:
        st.markdown(f"<div class='kpi'><div class='value'>{overall_ctr:.2f}%</div><div class='label'>📈 Overall CTR</div></div>", unsafe_allow_html=True)

    with col5:
        st.markdown(f"<div class='kpi'><div class='value'>{overall_cr:.2f}%</div><div class='label'>💡 Overall CR</div></div>", unsafe_allow_html=True)

    # ----------------- Chart Visualization -----------------
    st.write("### Performance Snapshot Metrics")
    # Chart to visualize Total Counts, Total Clicks, and Total Conversions
    try:
        # Prepare data for the chart
        chart_data = pd.DataFrame({
            'Metric': ['Total Counts', 'Total Clicks', 'Total Conversions'],
            'Counts': [total_counts, total_clicks, total_conv_safe]
        })
        # Calculate share as percentage of total_counts
        chart_data['Percentage'] = (chart_data['Counts'] / total_counts * 100).round(1)

        # Create a beautiful bar chart with text labels
        fig = px.bar(chart_data, x='Metric', y='Counts',
                    title='<b style="color:#FF5A6E; font-size:18px; text-shadow: 2px 2px 4px #00000055;">Performance Snapshot: 2025 Trends at a Glance! 🌟</b>',
                    labels={'Metric': '<i>Metric</i>', 'Counts': '<b>Search Counts</b>'},
                    color='Counts',
                    color_continuous_scale=['#E6F3FA', '#FFB085', '#FF5A6E'],
                    template='plotly_white',
                    text=chart_data['Counts'].astype(str))  # Show counts on bars

        # Update traces to position text outside, add percentages, and set hovertemplate
        fig.update_traces(
            texttemplate='%{text}<br>%{customdata:.1f}%',
            customdata=chart_data['Percentage'],
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Counts: %{y:,.0f}<br>Share: %{customdata:.1f}%<extra></extra>'
        )

        # Enhance attractiveness: Custom layout for beauty
        fig.update_layout(
            plot_bgcolor='rgba(255,255,255,0.95)',
            paper_bgcolor='rgba(255,247,232,0.8)',
            font=dict(color='#0B486B', family='Segoe UI'),
            title_x=0,  # Left alignment for title
            title_font_size=16,
            xaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
            yaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
            bargap=0.2,
            barcornerradius=8,
            annotations=[
                dict(
                    x=0.5, y=1.05, xref='paper', yref='paper',
                    text='✨ Hover for details | Peak metric highlighted below ✨',
                    showarrow=False,
                    font=dict(size=10, color='#FF5A6E', family='Segoe UI'),
                    align='center'
                )
            ]
        )

        # Highlight the peak metric with a custom marker or annotation
        peak_metric = chart_data.loc[chart_data['Counts'].idxmax(), 'Metric']
        peak_value = chart_data['Counts'].max()
        fig.add_annotation(
            x=peak_metric, y=peak_value,
            text=f"🏆 Peak: {peak_value:,.0f}",
            showarrow=True,
            arrowhead=3,
            arrowcolor='#FF5A6E',
            ax=0, ay=-30,
            font=dict(size=12, color='#FF5A6E', family='Segoe UI', weight='bold')
        )

        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error generating chart: {e}")

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
            # Check which columns actually exist before using them
            available_columns = queries.columns.tolist()
            agg_dict = {}
            
            if 'Counts' in available_columns:
                agg_dict['Counts'] = 'sum'
            if 'clicks' in available_columns:
                agg_dict['clicks'] = 'sum'
            if 'Conversion Rate' in available_columns:
                agg_dict['Conversion Rate'] = 'mean'
            
            # Only proceed if we have at least one column to aggregate
            if agg_dict:
                brand_perf = queries[queries['Brand'] != 'Other'].groupby('Brand').agg(agg_dict).reset_index()
                
                # Calculate derived metrics only if the required columns exist
                if 'clicks' in brand_perf.columns and 'Conversion Rate' in brand_perf.columns:
                    brand_perf['conversions'] = (brand_perf['clicks'] * brand_perf['Conversion Rate']).round()
                
                if 'Counts' in brand_perf.columns:
                    brand_perf['share'] = (brand_perf['Counts'] / total_counts * 100).round(2)
                
                # Only create the chart if we have data to display
                if not brand_perf.empty and 'Counts' in brand_perf.columns:
                    # Determine color column - use conversions if available, otherwise use Counts
                    color_column = 'conversions' if 'conversions' in brand_perf.columns else 'Counts'
                    hover_columns = ['share'] if 'share' in brand_perf.columns else []
                    if 'conversions' in brand_perf.columns:
                        hover_columns.append('conversions')
                    
                    # Create a beautiful bar chart with text labels
                    fig = px.bar(brand_perf.sort_values('Counts', ascending=False).head(10), 
                                x='Brand', y='Counts',
                                title='<b style="color:#FF5A6E; font-size:18px; text-shadow: 2px 2px 4px #00000055;">Top Brands by Search Counts</b>',
                                labels={'Brand': '<i>Brand</i>', 'Counts': '<b>Search Counts</b>'},
                                color=color_column,
                                color_continuous_scale=['#E6F3FA', '#FFB085', '#FF5A6E'],
                                template='plotly_white',
                                hover_data=hover_columns)
                    
                    # Update traces to position text outside and set hovertemplate
                    fig.update_traces(
                        texttemplate='%{y:,.0f}',
                        textposition='outside',
                        hovertemplate='<b>%{x}</b><br>Counts: %{y:,.0f}' + 
                                    ('<br>Share: %{customdata[0]:.2f}%' if 'share' in hover_columns else '') +
                                    ('<br>Conversions: %{customdata[1]:,.0f}' if 'conversions' in hover_columns and len(hover_columns) > 1 else '') +
                                    '<extra></extra>'
                    )

                    # Enhance attractiveness: Custom layout for beauty
                    fig.update_layout(
                        plot_bgcolor='rgba(255,255,255,0.95)',
                        paper_bgcolor='rgba(255,247,232,0.8)',
                        font=dict(color='#0B486B', family='Segoe UI'),
                        title_x=0,  # Left alignment for title
                        title_font_size=16,
                        xaxis=dict(
                            title='Brand',
                            showgrid=True, 
                            gridcolor='#E6F3FA', 
                            linecolor='#FF5A6E', 
                            linewidth=2
                        ),
                        yaxis=dict(
                            title='Search Counts',
                            showgrid=True, 
                            gridcolor='#E6F3FA', 
                            linecolor='#FF5A6E', 
                            linewidth=2
                        ),
                        bargap=0.2,
                        barcornerradius=8,
                        hovermode='x unified',
                        annotations=[
                            dict(
                                x=0.5, y=1.05, xref='paper', yref='paper',
                                text='✨ Hover for details | Top brand highlighted below ✨',
                                showarrow=False,
                                font=dict(size=10, color='#FF5A6E', family='Segoe UI'),
                                align='center'
                            )
                        ]
                    )

                    # Highlight the top brand with a custom marker
                    top_brand = brand_perf.loc[brand_perf['Counts'].idxmax(), 'Brand']
                    top_count = brand_perf['Counts'].max()
                    fig.add_annotation(
                        x=top_brand, y=top_count,
                        text=f"🏆 Peak: {top_count:,.0f}",
                        showarrow=True,
                        arrowhead=3,
                        arrowcolor='#FF5A6E',
                        ax=0, ay=-30,
                        font=dict(size=12, color='#FF5A6E', family='Segoe UI', weight='bold')
                    )

                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No brand data available after filtering or missing required columns.")
            else:
                st.warning("No valid aggregation columns found for brand analysis.")
        else:
            st.info("🏷 Brand column not found in the dataset.")

    with g2:
        if 'Category' in queries.columns:
            # Check which columns actually exist before using them
            available_columns = queries.columns.tolist()
            agg_dict = {}
            
            if 'Counts' in available_columns:
                agg_dict['Counts'] = 'sum'
            if 'clicks' in available_columns:
                agg_dict['clicks'] = 'sum'
            if 'conversions' in available_columns:
                agg_dict['conversions'] = 'sum'
            elif 'Conversion Rate' in available_columns and 'clicks' in available_columns:
                # We'll calculate conversions after aggregation
                pass
            
            # Only proceed if we have at least one column to aggregate
            if agg_dict:
                cat_perf = queries.groupby('Category').agg(agg_dict).reset_index()
                
                # Calculate conversions if we have the necessary columns but not the conversions column
                if 'conversions' not in cat_perf.columns and 'clicks' in cat_perf.columns and 'Conversion Rate' in queries.columns:
                    # Calculate average conversion rate for each category first
                    conv_rate_agg = queries.groupby('Category')['Conversion Rate'].mean().reset_index()
                    cat_perf = cat_perf.merge(conv_rate_agg, on='Category')
                    cat_perf['conversions'] = (cat_perf['clicks'] * cat_perf['Conversion Rate']).round()
                
                # Calculate share and conversion rate
                if 'Counts' in cat_perf.columns:
                    cat_perf['share'] = (cat_perf['Counts'] / total_counts * 100).round(2)
                
                # FIX: Calculate conversion rate correctly - conversions divided by counts
                if 'conversions' in cat_perf.columns and 'Counts' in cat_perf.columns:
                    cat_perf['cr'] = (cat_perf['conversions'] / cat_perf['Counts'] * 100).round(2)
                else:
                    cat_perf['cr'] = 0
                
                st.markdown("**Top Categories by Counts**")
                
                # Prepare display columns based on what's available
                display_columns = ['Category']
                format_dict = {}
                
                if 'Counts' in cat_perf.columns:
                    display_columns.append('Counts')
                    format_dict['Counts'] = '{:,.0f}'
                if 'share' in cat_perf.columns:
                    display_columns.append('share')
                    format_dict['share'] = '{:.2f}%'
                if 'clicks' in cat_perf.columns:
                    display_columns.append('clicks')
                    format_dict['clicks'] = '{:,.0f}'
                if 'conversions' in cat_perf.columns:
                    display_columns.append('conversions')
                    format_dict['conversions'] = '{:,.0f}'
                if 'cr' in cat_perf.columns:
                    display_columns.append('cr')
                    format_dict['cr'] = '{:.2f}%'
                
                # Display the table with available data
                if len(display_columns) > 1:  # More than just the Category column
                    try:
                        # Try using AgGrid if available
                        if 'AGGRID_OK' in globals() and AGGRID_OK:
                            AgGrid(cat_perf[display_columns].sort_values('Counts' if 'Counts' in display_columns else display_columns[1], ascending=False).head(10), 
                                height=300, enable_enterprise_modules=False)
                        else:
                            # Fall back to styled DataFrame
                            styled_cat_perf = cat_perf[display_columns].head(10).style.format(format_dict).set_properties(**{
                                'text-align': 'center',
                                'font-size': '14px'
                            }).background_gradient(subset=['cr'], cmap='YlGnBu')
                            st.dataframe(styled_cat_perf, use_container_width=True)
                    except NameError:
                        # AGGRID_OK not defined, use regular DataFrame
                        styled_cat_perf = cat_perf[display_columns].head(10).style.format(format_dict).set_properties(**{
                            'text-align': 'center',
                            'font-size': '14px'
                        }).background_gradient(subset=['cr'], cmap='YlGnBu')
                        st.dataframe(styled_cat_perf, use_container_width=True)
                else:
                    st.info("Insufficient data columns available for category analysis.")
            else:
                st.warning("No valid aggregation columns found for category analysis.")
        else:
            st.info("📦 Category column not found in the dataset.")

# ----------------- Search Analysis (Enhanced Core) -----------------
with tab_search:
    st.header("🔍 Search Column — Deep Dive Analysis")
    st.markdown("Analyze raw search queries with advanced keyword insights, performance metrics, and actionable intelligence. 🎯")
    
    # Hero Image for Search Tab
    search_image_options = {
        "Search Analytics Focus": "https://placehold.co/1200x200/FF5A6E/FFFFFF?text=🔍+Search+Query+Intelligence",
        "Data Visualization": "https://placehold.co/1200x200/E6F3FA/FF5A6E?text=📊+Keyword+Performance+Hub",
        "Abstract Search": "https://source.unsplash.com/1200x200/?analytics,data",
    }
    selected_search_image = st.sidebar.selectbox("Choose Search Tab Hero", options=list(search_image_options.keys()), index=0, key="search_hero")
    st.image(search_image_options[selected_search_image], use_container_width=True)
    
    # Add error handling and data validation
    if queries.empty or 'keywords' not in queries.columns:
        st.error("❌ No keyword data available. Please ensure your data contains properly processed keywords.")
        st.stop()
    
    # Quick Search Metrics Row
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        unique_queries = queries['normalized_query'].nunique()
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>🔍</span>
            <div class='value'>{unique_queries:,}</div>
            <div class='label'>Unique Search Queries</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_m2:
        avg_query_length = queries['query_length'].mean()
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>📏</span>
            <div class='value'>{avg_query_length:.1f}</div>
            <div class='label'>Avg Query Length</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_m3:
        total_keywords = sum(len(kw) for kw in queries['keywords'] if isinstance(kw, list))
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>🔤</span>
            <div class='value'>{total_keywords:,}</div>
            <div class='label'>Total Keywords</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_m4:
        long_tail_pct = (queries['query_length'] >= 20).mean() * 100
        st.markdown(f"""
        <div class='mini-metric'>
            <span class='icon'>📈</span>
            <div class='value'>{long_tail_pct:.1f}%</div>
            <div class='label'>Long-tail Queries</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Two-column layout for main analysis
    col_left, col_right = st.columns([3, 2])
    
    with col_left:
        # Enhanced Keyword Analysis
        st.subheader("🔤 Keyword Frequency & Performance Analysis")
        
        # Process keywords safely
        kw_series = queries['keywords'].explode().dropna()
        if kw_series.empty:
            st.warning("No keywords found in the dataset.")
        else:
            kw_counts = kw_series.value_counts().reset_index()
            kw_counts.columns = ['keyword', 'frequency']
            
            # Create keyword performance data
            keyword_performance = []
            for keyword in kw_counts['keyword'].head(50):  # Top 50 keywords
                keyword_queries = queries[queries['keywords'].apply(lambda x: keyword in x if isinstance(x, list) else False)]
                if not keyword_queries.empty:
                    performance = {
                        'keyword': keyword,
                        'frequency': len(keyword_queries),
                        'total_counts': keyword_queries['Counts'].sum(),
                        'total_clicks': keyword_queries['clicks'].sum(),
                        'total_conversions': keyword_queries['conversions'].sum(),
                        'avg_ctr': (keyword_queries['clicks'].sum() / keyword_queries['Counts'].sum() * 100) if keyword_queries['Counts'].sum() > 0 else 0,
                        'avg_cr': (keyword_queries['conversions'].sum() / keyword_queries['clicks'].sum() * 100) if keyword_queries['clicks'].sum() > 0 else 0
                    }
                    keyword_performance.append(performance)
            
            kw_perf_df = pd.DataFrame(keyword_performance)
            
            if not kw_perf_df.empty:
                # Enhanced keyword visualization
                fig_kw = px.scatter(
                    kw_perf_df.head(30), 
                    x='frequency', 
                    y='avg_ctr',
                    size='total_counts',
                    color='avg_cr',
                    hover_name='keyword',
                    title='<b style="color:#FF5A6E; font-size:18px;">Keyword Performance Matrix: Frequency vs CTR 🎯</b>',
                    labels={'frequency': 'Keyword Frequency', 'avg_ctr': 'Average CTR (%)', 'avg_cr': 'Avg CR (%)'},
                    color_continuous_scale=['#E6F3FA', '#FFB085', '#FF5A6E'],
                    template='plotly_white'
                )
                
                fig_kw.update_traces(
                    hovertemplate='<b>%{hovertext}</b><br>' +
                                 'Frequency: %{x}<br>' +
                                 'CTR: %{y:.2f}%<br>' +
                                 'Total Counts: %{marker.size:,.0f}<br>' +
                                 'Conversion Rate: %{marker.color:.2f}%<extra></extra>'
                )
                
                fig_kw.update_layout(
                    plot_bgcolor='rgba(255,255,255,0.95)',
                    paper_bgcolor='rgba(255,247,232,0.8)',
                    font=dict(color='#0B486B', family='Segoe UI'),
                    title_x=0,
                    xaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
                    yaxis=dict(showgrid=True, gridcolor='#E6F3FA', linecolor='#FF5A6E', linewidth=2),
                    annotations=[
                        dict(
                            x=0.95, y=0.95, xref='paper', yref='paper',
                            text='💡 Size = Total Counts | Color = Conversion Rate',
                            showarrow=False,
                            font=dict(size=11, color='#0B486B'),
                            align='right'
                        )
                    ]
                )
                
                st.plotly_chart(fig_kw, use_container_width=True)
                
                # Top performing keywords table
                st.subheader("🏆 Top Performing Keywords")
                top_keywords = kw_perf_df.sort_values('total_counts', ascending=False).head(15)

                # Create a copy for styling (keep numeric values)
                styling_kw = top_keywords.copy()

                # Format the dataframe for display AFTER applying styling
                display_kw = top_keywords.copy()
                display_kw['total_counts'] = display_kw['total_counts'].apply(lambda x: f"{x:,.0f}")
                display_kw['total_clicks'] = display_kw['total_clicks'].apply(lambda x: f"{x:,.0f}")
                display_kw['total_conversions'] = display_kw['total_conversions'].apply(lambda x: f"{x:,.0f}")
                display_kw['avg_ctr'] = display_kw['avg_ctr'].apply(lambda x: f"{x:.2f}%")
                display_kw['avg_cr'] = display_kw['avg_cr'].apply(lambda x: f"{x:.2f}%")

                display_kw = display_kw.rename(columns={
                    'keyword': 'Keyword',
                    'frequency': 'Frequency',
                    'total_counts': 'Total Counts',
                    'total_clicks': 'Total Clicks',
                    'total_conversions': 'Conversions',
                    'avg_ctr': 'Avg CTR',
                    'avg_cr': 'Avg CR'
                })

                # Apply styling to the numeric version, then format
                styled_keywords = styling_kw.style.background_gradient(
                    subset=['frequency'], 
                    cmap='Blues', 
                    alpha=0.3
                ).format({
                    'total_counts': '{:,.0f}',
                    'total_clicks': '{:,.0f}',
                    'total_conversions': '{:,.0f}',
                    'avg_ctr': '{:.2f}%',
                    'avg_cr': '{:.2f}%'
                }).set_properties(**{
                    'text-align': 'center',
                    'font-size': '13px'
                })

                # Rename columns for final display
                styled_keywords = styled_keywords.relabel_index({
                    'keyword': 'Keyword',
                    'frequency': 'Frequency',
                    'total_counts': 'Total Counts',
                    'total_clicks': 'Total Clicks',
                    'total_conversions': 'Conversions',
                    'avg_ctr': 'Avg CTR',
                    'avg_cr': 'Avg CR'
                }, axis=1)

                st.dataframe(styled_keywords, use_container_width=True)

    
    with col_right:
        # Word Cloud (Enhanced)
        if WORDCLOUD_OK and not kw_counts.empty:
            st.subheader("🌟 Keyword Cloud Visualization")
            try:
                freqs = dict(zip(kw_counts['keyword'].head(100), kw_counts['frequency'].head(100)))
                wc = WordCloud(
                    width=500, 
                    height=300, 
                    background_color='rgba(255,247,232,0.8)',
                    colormap='plasma',
                    collocations=False, 
                    max_words=80,
                    relative_scaling=0.5,
                    min_font_size=8
                ).generate_from_frequencies(freqs)
                
                fig_wc, ax_wc = plt.subplots(figsize=(8, 5))
                ax_wc.imshow(wc, interpolation='bilinear')
                ax_wc.axis('off')
                ax_wc.set_facecolor('rgba(255,247,232,0.8)')
                fig_wc.patch.set_facecolor('rgba(255,247,232,0.8)')
                st.pyplot(fig_wc, use_container_width=True)
                plt.close(fig_wc)
            except Exception as e:
                st.error(f"Error generating word cloud: {e}")
        else:
            if not WORDCLOUD_OK:
                st.info("💡 Install 'wordcloud' and 'matplotlib' for enhanced visualizations")
        
        # Query Length Distribution
        st.subheader("📊 Query Length Analysis")
        length_dist = queries.groupby('query_length').size().reset_index(name='count')
        length_dist = length_dist.sort_values('query_length')
        
        fig_length = px.histogram(
            queries, 
            x='query_length', 
            nbins=30,
            title='<b style="color:#FF5A6E;">Query Length Distribution</b>',
            labels={'query_length': 'Character Length', 'count': 'Number of Queries'},
            color_discrete_sequence=['#FF8A7A']
        )
        
        fig_length.update_layout(
            plot_bgcolor='rgba(255,255,255,0.95)',
            paper_bgcolor='rgba(255,247,232,0.8)',
            font=dict(color='#0B486B', family='Segoe UI'),
            bargap=0.1,
            xaxis=dict(showgrid=True, gridcolor='#E6F3FA'),
            yaxis=dict(showgrid=True, gridcolor='#E6F3FA')
        )
        
        st.plotly_chart(fig_length, use_container_width=True)
    
    st.markdown("---")
    
    # Advanced Analytics Section
    st.subheader("📈 Advanced Query Performance Analytics")
    
    # Three-column layout for advanced metrics
    adv_col1, adv_col2, adv_col3 = st.columns(3)
    
    with adv_col1:
        st.markdown("**🎯 Query Length vs Performance**")
        ql_analysis = queries.groupby('query_length').agg({
            'Counts': 'sum', 
            'clicks': 'sum',
            'conversions': 'sum'
        }).reset_index()
        ql_analysis['ctr'] = ql_analysis.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        ql_analysis['cr'] = ql_analysis.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        
        if not ql_analysis.empty:
            fig_ql = px.scatter(
                ql_analysis, 
                x='query_length', 
                y='ctr', 
                size='Counts',
                color='cr',
                title='Length vs CTR Performance',
                color_continuous_scale=['#E6F3FA', '#FF8A7A'],
                template='plotly_white'
            )
            
            fig_ql.update_layout(
                plot_bgcolor='rgba(255,255,255,0.95)',
                paper_bgcolor='rgba(255,247,232,0.8)',
                font=dict(color='#0B486B', family='Segoe UI', size=10),
                height=300,
                xaxis=dict(showgrid=True, gridcolor='#E6F3FA'),
                yaxis=dict(showgrid=True, gridcolor='#E6F3FA')
            )
            
            st.plotly_chart(fig_ql, use_container_width=True)
    
    with adv_col2:
        st.markdown("**📊 Long-tail vs Short-tail Performance**")
        queries['is_long_tail'] = queries['query_length'] >= 20
        lt_analysis = queries.groupby('is_long_tail').agg({
            'Counts': 'sum', 
            'clicks': 'sum',
            'conversions': 'sum'
        }).reset_index()
        lt_analysis['label'] = lt_analysis['is_long_tail'].map({
            True: 'Long-tail (≥20 chars)', 
            False: 'Short-tail (<20 chars)'
        })
        lt_analysis['ctr'] = lt_analysis.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        
        if not lt_analysis.empty:
            fig_lt = px.bar(
                lt_analysis, 
                x='label', 
                y='Counts',
                color='ctr',
                title='Traffic: Long-tail vs Short-tail',
                color_continuous_scale=['#E6F3FA', '#FF5A6E'],
                text='Counts'
            )
            
            fig_lt.update_traces(
                texttemplate='%{text:,.0f}',
                textposition='outside'
            )
            
            fig_lt.update_layout(
                plot_bgcolor='rgba(255,255,255,0.95)',
                paper_bgcolor='rgba(255,247,232,0.8)',
                font=dict(color='#0B486B', family='Segoe UI', size=10),
                height=300,
                xaxis=dict(showgrid=True, gridcolor='#E6F3FA'),
                yaxis=dict(showgrid=True, gridcolor='#E6F3FA')
            )
            
            st.plotly_chart(fig_lt, use_container_width=True)
    
    with adv_col3:
        st.markdown("**🔍 Search Intent Analysis**")
        # Create search intent categories based on keywords
        intent_mapping = {
            'buy': ['buy', 'purchase', 'order', 'shop'],
            'compare': ['vs', 'versus', 'compare', 'difference', 'better'],
            'info': ['what', 'how', 'why', 'when', 'where'],
            'brand': ['brand', 'reviews', 'rating', 'best']
        }
        
        def classify_intent(keywords_list):
            if not isinstance(keywords_list, list):
                return 'other'
            keywords_str = ' '.join(keywords_list).lower()
            for intent, intent_keywords in intent_mapping.items():
                if any(kw in keywords_str for kw in intent_keywords):
                    return intent
            return 'other'
        
        queries['search_intent'] = queries['keywords'].apply(classify_intent)
        intent_analysis = queries.groupby('search_intent').agg({
            'Counts': 'sum',
            'clicks': 'sum',
            'conversions': 'sum'
        }).reset_index()
        intent_analysis['ctr'] = intent_analysis.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
        
        if not intent_analysis.empty:
            fig_intent = px.pie(
                intent_analysis, 
                names='search_intent', 
                values='Counts',
                title='Search Intent Distribution',
                color_discrete_sequence=['#FF5A6E', '#FFB085', '#E6F3FA', '#FF8A7A', '#FFF7E8']
            )
            
            fig_intent.update_layout(
                font=dict(color='#0B486B', family='Segoe UI', size=10),
                height=300
            )
            
            st.plotly_chart(fig_intent, use_container_width=True)
    
    st.markdown("---")
    
    # Query Performance Insights Table
    st.subheader("📋 Detailed Query Performance Analysis")
    
    # Create comprehensive query analysis
    query_analysis = queries.groupby('normalized_query').agg({
        'Counts': 'sum',
        'clicks': 'sum', 
        'conversions': 'sum',
        'query_length': 'first'
    }).reset_index()
    
    query_analysis['ctr'] = query_analysis.apply(lambda r: (r['clicks']/r['Counts']*100) if r['Counts']>0 else 0, axis=1)
    query_analysis['cr'] = query_analysis.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
    query_analysis['is_long_tail'] = query_analysis['query_length'] >= 20
    query_analysis['performance_score'] = (query_analysis['ctr'] * 0.6 + query_analysis['cr'] * 0.4).round(2)
    
    # Sort by performance score and get top performers
    top_performers = query_analysis.sort_values('performance_score', ascending=False).head(20)
    
    # Format for display
    display_analysis = top_performers.copy()
    display_analysis['Counts'] = display_analysis['Counts'].apply(lambda x: f"{x:,.0f}")
    display_analysis['clicks'] = display_analysis['clicks'].apply(lambda x: f"{x:,.0f}")
    display_analysis['conversions'] = display_analysis['conversions'].apply(lambda x: f"{x:,.0f}")
    display_analysis['ctr'] = display_analysis['ctr'].apply(lambda x: f"{x:.2f}%")
    display_analysis['cr'] = display_analysis['cr'].apply(lambda x: f"{x:.2f}%")
    display_analysis['is_long_tail'] = display_analysis['is_long_tail'].map({True: '✅', False: '❌'})
    
    display_analysis = display_analysis.rename(columns={
        'normalized_query': 'Search Query',
        'Counts': 'Total Searches',
        'clicks': 'Clicks',
        'conversions': 'Conversions',
        'ctr': 'CTR',
        'cr': 'CR',
        'query_length': 'Length',
        'is_long_tail': 'Long-tail',
        'performance_score': 'Score'
    })
    
    # Display with enhanced styling
    styled_analysis = display_analysis[['Search Query', 'Total Searches', 'Clicks', 'Conversions', 'CTR', 'CR', 'Length', 'Long-tail', 'Score']].style.format({
        'Score': '{:.2f}'
    }).set_properties(**{
        'text-align': 'center',
        'font-size': '12px'
    }).background_gradient(subset=['Score'], cmap='RdYlGn', alpha=0.4)
    
    st.dataframe(styled_analysis, use_container_width=True)
    
    # Download functionality
    csv_analysis = query_analysis.to_csv(index=False)
    st.download_button(
        label="📥 Download Query Analysis CSV",
        data=csv_analysis,
        file_name="query_performance_analysis.csv",
        mime="text/csv"
    )
    
    # Key Insights Box
    st.markdown("---")
    col_insight1, col_insight2 = st.columns(2)
    
    with col_insight1:
        st.markdown("""
        <div class='insight-box'>
            <h4>🎯 Key Insights</h4>
            <p>• Long-tail queries represent {:.1f}% of total traffic<br>
            • Average query length is {:.1f} characters<br>
            • Top keyword appears in {:.1f}% of searches</p>
        </div>
        """.format(
            long_tail_pct,
            avg_query_length,
            (kw_counts.iloc[0]['frequency'] / len(queries) * 100) if not kw_counts.empty else 0
        ), unsafe_allow_html=True)
    
    with col_insight2:
        st.markdown("""
        <div class='insight-box'>
            <h4>💡 Recommendations</h4>
            <p>• Focus on high-performing keywords for content optimization<br>
            • Analyze long-tail queries for niche opportunities<br>
            • Monitor search intent patterns for strategy alignment</p>
        </div>
        """, unsafe_allow_html=True)
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