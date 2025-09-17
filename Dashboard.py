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
# [Unchanged CSS from original code]
st.markdown("""
<style>
/* Same CSS as provided */
</style>
""", unsafe_allow_html=True)

# ----------------- Helpers -----------------
# [Unchanged helper functions: safe_read_excel, extract_keywords, prepare_queries_df]
# (These remain identical to the original code)

# ----------------- Data load (upload or default) -----------------
# [Unchanged data loading logic]
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
# [Unchanged sheet selection logic]
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
# [Unchanged filter logic]
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
# [Unchanged welcome message]
st.markdown("""
<div class="welcome-box">
    <h2>👋 Welcome to Lady Care Search Analytics! ✨</h2>
    <p>Explore search patterns, brand performance, and actionable insights. Use the sidebar to filter data, navigate tabs to dive deep, and download results for your reports!</p>
</div>
""", unsafe_allow_html=True)

# ----------------- KPI cards -----------------
# [Unchanged KPI cards]
st.markdown('<div class="main-header">🔥 Lady Care — Ultimate Search Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Uncover powerful insights from the <b>search</b> column with vibrant visuals and actionable pivots</div>', unsafe_allow_html=True)

total_impr = int(queries['impressions'].sum())
total_clicks = int(queries['clicks'].sum())
total_conv = int(queries['conversions'].sum())
overall_ctr = (queries['clicks'].sum()/queries['impressions'].sum()) * 100 if queries['impressions'].sum()>0 else 0
overall_cr = (queries['conversions'].sum()/queries['clicks'].sum()) * 100 if queries['clicks'].sum()>0 else 0

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

# ----------------- Overview (Modified) -----------------
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
    st.subheader("🏷 Brand & Category Snapshot")  # Replaced Geography & Device
    g1, g2 = st.columns(2)
    with g1:
        if 'brand' in queries.columns and queries['brand'].notna().any():
            brand_perf = queries.groupby('brand').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(10)
            fig = px.bar(brand_perf, x='brand', y='impressions', title='Top 10 Brands by Impressions', color_discrete_sequence=px.colors.qualitative.D3)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Brand column not available.")
    with g2:
        if 'category' in queries.columns and queries['category'].notna().any():
            cat_perf = queries.groupby('category').agg(conversions=('conversions','sum'), impressions=('impressions','sum')).reset_index().sort_values('conversions', ascending=False).head(10)
            st.markdown("**Top 10 Categories by Conversions**")
            if AGGRID_OK:
                gb = GridOptionsBuilder.from_dataframe(cat_perf)
                gb.configure_grid_options(enableRangeSelection=True)
                AgGrid(cat_perf, gridOptions=gb.build(), height=300)
            else:
                st.dataframe(cat_perf, use_container_width=True)
        else:
            st.info("Category column not available.")

# ----------------- Search Analysis (core) -----------------
# [Unchanged Search Analysis tab]

# ----------------- Brand Tab -----------------
# [Unchanged Brand tab]

# ----------------- Category Tab -----------------
# [Unchanged Category tab]

# ----------------- Subcategory Tab -----------------
# [Unchanged Subcategory tab]

# ----------------- Generic Type Tab -----------------
# [Unchanged Generic Type tab]

# ----------------- Time Analysis (Modified) -----------------
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

    # Replaced CTR by Day of Week with Brand Impressions by Month
    st.subheader("🏷 Top Brands by Month (Impressions)")
    if 'brand' in queries.columns and queries['brand'].notna().any() and queries['month'].notna().any():
        top_brands = queries.groupby('brand')['impressions'].sum().sort_values(ascending=False).head(5).index
        brand_month = queries[queries['brand'].isin(top_brands)].groupby(['month','brand']).agg(impressions=('impressions','sum')).reset_index()
        try:
            brand_month['month_dt'] = pd.to_datetime(brand_month['month'], format='%b %Y', errors='coerce')
            brand_month = brand_month.sort_values('month_dt')
        except:
            brand_month = brand_month.sort_values('month')
        fig = px.bar(brand_month, x='month', y='impressions', color='brand', title='Top 5 Brands by Impressions per Month', color_discrete_sequence=px.colors.qualitative.D3)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Brand or month data not available for brand-month analysis.")

# ----------------- Pivot Builder Tab -----------------
# [Unchanged Pivot Builder tab]

# ----------------- Insights & Actionable Questions (Modified) -----------------
with tab_insights:
    st.header("💡 Insights & Actionable Questions (25)")
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

    # Q3: Top queries by conversion rate (originally Q4)
    def q3():
        df4 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df4 = df4[df4['impressions']>=50]
        df4['cr'] = df4.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df4.sort_values('cr', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q3 — Top Queries by Conversion Rate (Min Impressions=50)", "High-converting queries for paid promotions.", q3, "🎯")

    # Q4: Long-tail contribution (originally Q5)
    def q4():
        lt = queries[queries['query_length']>=20]
        st.markdown(f"Long-tail rows: {len(lt):,} / total {len(queries):,}")
        st.plotly_chart(px.pie(names=['Long-tail','Rest'], values=[lt['impressions'].sum(), queries['impressions'].sum()-lt['impressions'].sum()], title='Impression Share: Long-Tail vs Rest', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
    q_expand("Q4 — Long-Tail vs Short-Tail (>=20 chars)", "How much traffic comes from long-tail queries? Key for content strategy.", q4, "📏")

    # Q5: Brand vs generic share (originally Q7)
    def q5():
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
    q_expand("Q5 — Branded vs Generic Queries (Pivot)", "Assess brand vs generic search intent with a pivot table.", q5, "🏷")

    # Q6: Rising queries MoM (originally Q8)
    def q6():
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
    q_expand("Q6 — Top Rising Queries Month-over-Month", "Detect emerging demand for seasonal campaigns.", q6, "📈")

    # Q7: Query funnel snapshot (originally Q11)
    def q7():
        snap = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(200)
        if AGGRID_OK:
            AgGrid(snap, height=400)
        else:
            st.dataframe(snap.head(100), use_container_width=True)
    q_expand("Q7 — Query Funnel Snapshot (Top 200)", "View top queries' funnel: impressions → clicks → conversions.", q7, "📊")

    # Q8: Traffic concentration (originally Q14)
    def q8():
        qq = queries.groupby('normalized_query').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions', ascending=False)
        top5n = max(1, int(0.05*len(qq)))
        share = qq.head(top5n)['impressions'].sum() / qq['impressions'].sum() if qq['impressions'].sum()>0 else 0
        st.markdown(f"Top 5% queries contribute **{share:.2%}** of impressions (top {top5n} queries).")
        st.plotly_chart(px.pie(names=['Top 5% Queries','Rest'], values=[qq.head(top5n)['impressions'].sum(), qq['impressions'].sum()-qq.head(top5n)['impressions'].sum()], title='Traffic Concentration: Top 5% Queries', color_discrete_sequence=px.colors.qualitative.D3), use_container_width=True)
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
        dfm = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        out = dfm[(dfm['impressions']>=dfm['impressions'].quantile(0.7)) & (dfm['conversions']==0)].sort_values('impressions', ascending=False).head(40)
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
        df19 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum')).reset_index()
        df19 = df19[df19['impressions']>=30]
        df19['ctr'] = df19.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        out = df19.sort_values('ctr', ascending=False).head(40)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q12 — Top Queries by CTR (Min Impressions=30)", "High-engagement queries for ad campaigns.", q12, "📈")

    # Q13: Low CTR & CR queries (originally Q20)
    def q13():
        df20 = queries.groupby('normalized_query').agg(impressions=('impressions','sum'), clicks=('clicks','sum'), conversions=('conversions','sum')).reset_index()
        df20['ctr'] = df20.apply(lambda r: (r['clicks']/r['impressions']*100) if r['impressions']>0 else 0, axis=1)
        df20['cr'] = df20.apply(lambda r: (r['conversions']/r['clicks']*100) if r['clicks']>0 else 0, axis=1)
        out = df20[(df20['impressions']>=df20['impressions'].quantile(0.6)) & (df20['ctr']<=df20['ctr'].quantile(0.25)) & (df20['cr']<=df20['cr'].quantile(0.25))].sort_values('impressions', ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q13 — High Impressions, Low CTR & CR", "Optimize search results for these underperforming queries.", q13, "⚠️")

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
        out = queries.groupby('normalized_query').agg(conversions=('conversions','sum'), impressions=('impressions','sum')).reset_index().sort_values('conversions', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q16 — Top Queries by Conversions", "Direct revenue drivers for inventory and bids.", q16, "🎯")

    # Q17: Month-by-month query trends (originally Q27)
    def q17():
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
        m1 = queries[queries['month']==first].groupby('normalized_query').agg(impressions=('impressions','sum')).rename(columns={'impressions':'m1'})
        m2 = queries[queries['month']==last].groupby('normalized_query').agg(impressions=('impressions','sum')).rename(columns={'impressions':'m2'})
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
        out = queries.groupby('normalized_query').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions',ascending=False).head(50)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q20 — Top 50 Queries (Quick)", "Quick ranking of top queries by impressions.", q20, "📊")

    # Q21: Top brands quick view (originally Q32)
    def q21():
        if 'brand' in queries.columns:
            out = queries.groupby('brand').agg(impressions=('impressions','sum'), conversions=('conversions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Brand missing.")
    q_expand("Q21 — Top Brands (Quick)", "Quick brand ranking by impressions.", q21, "🏷")

    # Q22: Top subcategories quick view (originally Q33)
    def q22():
        if 'sub_category' in queries.columns:
            out = queries.groupby('sub_category').agg(impressions=('impressions','sum')).reset_index().sort_values('impressions', ascending=False).head(50)
            if AGGRID_OK:
                AgGrid(out, height=400)
            else:
                st.dataframe(out, use_container_width=True)
        else:
            st.info("Subcategory missing.")
    q_expand("Q22 — Top Subcategories (Quick)", "Quick subcategory ranking.", q22, "🧴")

    # Q23: Top queries by clicks (originally Q37)
    def q23():
        out = queries.groupby('normalized_query').agg(clicks=('clicks','sum'), impressions=('impressions','sum')).reset_index().sort_values('clicks', ascending=False).head(30)
        if AGGRID_OK:
            AgGrid(out, height=400)
        else:
            st.dataframe(out, use_container_width=True)
    q_expand("Q23 — Top Queries by Clicks", "High-engagement queries for ad optimization.", q23, "👆")

    # Q24: Category vs brand performance (originally Q38)
    def q24():
        if 'category' in queries.columns and 'brand' in queries.columns:
            pivot = queries.pivot_table(values=['impressions','clicks','conversions'], index=['category'], columns=['brand'], aggfunc='sum').fillna(0)
            if AGGRID_OK:
                AgGrid(pivot.reset_index(), height=400)
            else:
                st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Category or brand missing.")
    q_expand("Q24 — Category vs Brand Performance (Pivot)", "Analyze brand performance within categories.", q24, "📦🏷")

    # Q25: High impressions, low clicks by category (originally Q39)
    def q25():
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
    q_expand("Q25 — High Impressions, Low Clicks by Category", "Identify category-specific queries needing optimization.", q25, "⚠️")

    st.info("Want more advanced questions (e.g., anomaly detection, semantic clustering)? I can add them with additional packages like scikit-learn or prophet.")

# ----------------- Export / Downloads -----------------
# [Unchanged Export tab]

# ----------------- Footer -----------------
# [Unchanged Footer]
st.markdown(f"""
<div class="footer">
✨ Lady Care Search Analytics — Noureldeen Mohamed
</div>
""", unsafe_allow_html=True)