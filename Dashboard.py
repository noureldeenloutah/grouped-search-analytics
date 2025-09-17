
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import re, os, logging
from collections import Counter
from datetime import datetime

# Optional interactive grid
try:
    from st_aggrid import AgGrid, GridOptionsBuilder
    AGGRID_OK = True
except Exception:
    AGGRID_OK = False

# Optional wordcloud
try:
    from wordcloud import WordCloud
    WORDCLOUD_OK = True
except Exception:
    WORDCLOUD_OK = False

# Page config & style
st.set_page_config(page_title="🔥 Lady Care — Search Query Intelligence", layout="wide", page_icon="🔎")
st.markdown("""
<style>
.kpi {background:#E8F6F8; padding:12px; border-radius:8px; text-align:center}
.small {font-size:0.9rem; color:#606c76}
.big {font-size:1.6rem; font-weight:700; color:#023047}
.card {background: linear-gradient(90deg,#E6F6F8,#F7FDFF); padding:12px; border-radius:10px}
</style>
""", unsafe_allow_html=True)

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ----------------- Helpers -----------------
@st.cache_data(show_spinner=False)
def load_data(path="Lady Care Preprocessed Data.xlsx"):
    """
    Load preprocessed dataset. Adjust path as needed.
    Expected columns: Date, normalized_query, CTR, CR, Count, Rev, Month, day_of_week, week_number,
    language, processed_query, misspelling_type, match_confidence, conversions, conversion_value, revenue_per_search
    """
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, parse_dates=['Date'], dayfirst=True)
    else:
        df = pd.read_excel(path, parse_dates=['Date'])
    return df

def safe_lower(x):
    try: return str(x).lower()
    except: return ""

@st.cache_data
def preprocess(df):
    df = df.copy()
    # Keep original search text (do NOT autocorrect)
    for c in ["normalized_query", "processed_query"]:
        if c not in df.columns:
            df[c] = ""
    # computed
    df["query_length"] = df["normalized_query"].astype(str).apply(len)
    # percentages - ensure in 0..100
    for p in ["CTR","CR","ATCR"]:
        if p in df.columns:
            df[p] = pd.to_numeric(df[p], errors="coerce") * 100
        else:
            df[p] = 0.0
    # ensure numeric
    for n in ["Count","Rev","conversions","conversion_value","revenue_per_search"]:
        if n in df.columns:
            df[n] = pd.to_numeric(df[n], errors="coerce").fillna(0)
        else:
            df[n] = 0
    # Safe month/year
    if "Month" not in df.columns:
        df["Month"] = pd.to_datetime(df["Date"]).dt.strftime("%b")
    if "Year" not in df.columns:
        df["Year"] = pd.to_datetime(df["Date"]).dt.year
    df["Month_Year"] = pd.to_datetime(df["Date"]).dt.to_period("M").astype(str)
    # CTR bucket
    df["CTR_bucket"] = pd.qcut(df["CTR"].rank(method="first"), q=4, labels=["Low","Med","High","Very High"])
    return df

def top_n_queries(df, n=20):
    return df['normalized_query'].value_counts().head(n).reset_index().rename(columns={"index":"query", "normalized_query":"count"})

def most_common_words(df, col="normalized_query", top_n=30):
    all_text = " ".join(df[col].astype(str).tolist()).lower()
    words = re.findall(r'\w+', all_text)
    common = Counter(words).most_common(top_n)
    return pd.DataFrame(common, columns=["word","count"])

def build_pivot(df, index, columns, values, aggfunc="sum"):
    try:
        pt = pd.pivot_table(df, index=index, columns=columns, values=values, aggfunc=aggfunc, fill_value=0)
        pt = pt.reset_index()
        return pt
    except Exception as e:
        logger.error("Pivot error: %s", e)
        return pd.DataFrame()

def kpi_row(df):
    imps = int(df["Count"].sum())
    clicks = int((df["Count"] * (df["CTR"]/100)).sum() if "CTR" in df.columns else 0)
    conv = int(df.get("conversions", pd.Series(0)).sum())
    rev = float(df.get("Rev", pd.Series(0)).sum())
    return imps, clicks, conv, rev

def aggrid_display(df, height=350):
    if not AGGRID_OK:
        st.dataframe(df)
        return
    builder = GridOptionsBuilder.from_dataframe(df)
    builder.configure_default_column(filterable=True, sortable=True, resizable=True)
    builder.configure_selection(selection_mode="single", use_checkbox=False)
    gridOptions = builder.build()
    AgGrid(df, gridOptions=gridOptions, enable_enterprise_modules=False, height=height, fit_columns_on_grid_load=True)

# ----------------- Load & Prep -----------------
st.sidebar.title("🔎 Filters & Controls")
data_path = st.sidebar.text_input("Dataset path (xlsx/csv)", value="Lady Care Preprocessed Data.xlsx")
sample_size = st.sidebar.slider("Sample size for visuals", min_value=5000, max_value=200000, value=50000, step=5000)
run_wordcloud = st.sidebar.checkbox("Show word cloud", value=True)
show_aggrid = st.sidebar.checkbox("Use interactive pivot (AgGrid)", value=True if AGGRID_OK else False)

df_raw = load_data(data_path)
if df_raw is None:
    st.error("Couldn't load dataset. Check path or file.")
    st.stop()

df = preprocess(df_raw)

# Sidebar filters UI
languages = ["All"] + sorted(df["language"].dropna().unique().tolist())
selected_lang = st.sidebar.multiselect("Select language(s)", languages, default=["All"])
subcats = ["All"] + sorted(df["Sub Category"].dropna().unique().tolist()) if "Sub Category" in df.columns else ["All"]
selected_subcat = st.sidebar.multiselect("Select subcategory(ies)", subcats, default=["All"])
months = ["All"] + sorted(df["Month_Year"].unique().tolist())
selected_months = st.sidebar.multiselect("Select Month(s)", months, default=["All"])
text_search = st.sidebar.text_input("Filter by query text (contains)", value="")

# Apply filters
df_filtered = df.copy()
if "All" not in selected_lang:
    df_filtered = df_filtered[df_filtered["language"].isin(selected_lang)]
if "All" not in selected_subcat and "Sub Category" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Sub Category"].isin(selected_subcat)]
if "All" not in selected_months:
    df_filtered = df_filtered[df_filtered["Month_Year"].isin(selected_months)]
if text_search.strip():
    df_filtered = df_filtered[df_filtered["normalized_query"].str.contains(text_search, case=False, na=False)]
# sample for visual performance
df_sample = df_filtered.sample(n=min(sample_size, max(1, len(df_filtered))), random_state=42) if len(df_filtered)>sample_size else df_filtered

# ----------------- Top KPI row -----------------
st.markdown("## 🔥 Lady Care — Search Query Intelligence")
imps, clicks, conv, rev = kpi_row(df_filtered)
col1, col2, col3, col4 = st.columns([1.2,1,1,1])
col1.markdown(f"<div class='card'><div class='small'>Total Searches (impressions)</div><div class='big'>{imps:,}</div></div>", unsafe_allow_html=True)
col2.markdown(f"<div class='card'><div class='small'>Estimated Clicks</div><div class='big'>{clicks:,}</div></div>", unsafe_allow_html=True)
col3.markdown(f"<div class='card'><div class='small'>Conversions</div><div class='big'>{conv:,}</div></div>", unsafe_allow_html=True)
col4.markdown(f"<div class='card'><div class='small'>Revenue (SAR)</div><div class='big'>{rev:,.2f}</div></div>", unsafe_allow_html=True)

# ----------------- Tabs -----------------
tabs = st.tabs(["🏠 Overview", "📊 Univariate", "📈 Bivariate", "🔍 Categorical", "📅 Temporal", "💡 Insights Hub", "📋 Pivot & Export"])

# ---------- Overview Tab ----------
with tabs[0]:
    st.header("Overview — quick interactive snapshots")
    c1, c2 = st.columns([2,1])
    with c1:
        st.subheader("Top 10 Queries by Search Count")
        top10 = top_n_queries(df_filtered, 10)
        fig = px.bar(top10, x="query", y="count", title="Top 10 Most Frequent Search Queries", labels={"query":"Query","count":"Count"})
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("**Hint:** click a bar to copy the query text and inspect it on the Pivot tab.")
    with c2:
        st.subheader("Most Common Words (Top 15)")
        wc = most_common_words(df_sample, top_n=50)
        if WORDCLOUD_OK and run_wordcloud:
            text = " ".join([f"{w} " * int(c/ max(1, wc['count'].max()//50)) for w,c in wc.values])
            image = WordCloud(width=400, height=240, background_color="white").generate(text)
            st.image(image.to_array(), use_column_width=True)
        else:
            st.bar_chart(wc.head(15).set_index("word")["count"])

    st.markdown("---")
    st.subheader("Small sample preview of filtered data")
    st.dataframe(df_filtered.head(8), use_container_width=True)

# ---------- Univariate Tab ----------
with tabs[1]:
    st.header("Univariate analyses (single-dimension)")
    # 1 Top 20 queries by Count
    st.subheader("1) Top 20 Queries by Count")
    st.plotly_chart(px.bar(top_n_queries(df_filtered,20), x="query", y="count", title="Top 20 Queries (by Count)"), use_container_width=True)

    # 2 Distribution of Counts (log)
    st.subheader("2) Distribution of Search Count (log-scale)")
    st.plotly_chart(px.histogram(df_sample, x="Count", nbins=60, title="Search Count Distribution (log y)", log_y=True), use_container_width=True)

    # 3 CTR distribution
    st.subheader("3) CTR Distribution (%)")
    st.plotly_chart(px.histogram(df_sample, x="CTR", nbins=60, title="CTR Distribution (%)"), use_container_width=True)

    # 4 Query length distribution
    st.subheader("4) Query Length (characters)")
    st.plotly_chart(px.histogram(df_sample, x="query_length", nbins=40, title="Query Length Distribution"), use_container_width=True)

    # 5 Top 20 queries by revenue
    st.subheader("5) Top 20 Queries by Revenue")
    st.plotly_chart(px.bar(df_filtered.sort_values("Rev", ascending=False).head(20), x="normalized_query", y="Rev", title="Top 20 Queries by Revenue"), use_container_width=True)

# ---------- Bivariate Tab ----------
with tabs[2]:
    st.header("Bivariate analyses (relationships)")
    # 6 Count vs Revenue (heatmap)
    st.subheader("6) Count vs Revenue (Avg CTR color)")
    st.plotly_chart(px.density_heatmap(df_sample, x="Count", y="Rev", z="CTR", nbinsx=50, nbinsy=50, title="Count vs Revenue (color = avg CTR)"), use_container_width=True)

    # 7 CTR vs CR
    st.subheader("7) CTR vs Conversion Rate (CR)")
    st.plotly_chart(px.density_heatmap(df_sample, x="CTR", y="CR", z="Count", nbinsx=50, nbinsy=50, title="CTR vs CR (Count as color)"), use_container_width=True)

    # 8 Query length vs CTR scatter
    st.subheader("8) Query Length vs CTR (size = Count)")
    st.plotly_chart(px.scatter(df_sample, x="query_length", y="CTR", size="Count", color="CR", hover_data=["normalized_query"], title="Query Length vs CTR"), use_container_width=True)

    # 9 Revenue per Conversion vs Query Length
    if "conversions" in df.columns and df["conversions"].sum()>0:
        st.subheader("9) Revenue per Conversion by Query Length Bucket")
        df_temp = df_sample.copy()
        df_temp['len_bucket'] = pd.cut(df_temp['query_length'], bins=[0,8,16,32,200], labels=['Very Short','Short','Medium','Long'])
        rpc = df_temp.groupby('len_bucket').apply(lambda x: (x['Rev'].sum() / x['conversions'].sum()) if x['conversions'].sum()>0 else 0).reset_index(name='rev_per_conv')
        st.plotly_chart(px.bar(rpc, x='len_bucket', y='rev_per_conv', title="Revenue per Conversion by Query Length"), use_container_width=True)

# ---------- Categorical Tab ----------
with tabs[3]:
    st.header("Categorical analyses (language, brand, misspellings, subcategory)")
    # 10 Language distribution
    st.subheader("10) Language Distribution of Queries")
    st.plotly_chart(px.pie(df_filtered, names="language", title="Language Distribution"), use_container_width=True)

    # 11 CTR by language
    st.subheader("11) Average CTR by Language")
    st.plotly_chart(px.bar(df_filtered.groupby("language")["CTR"].mean().reset_index().sort_values("CTR", ascending=False), x="language", y="CTR", title="Avg CTR by Language"), use_container_width=True)

    # 12 Misspelling counts & impact
    if "misspelling_type" in df_filtered.columns:
        st.subheader("12) Misspelling Types & Avg CTR")
        m = df_filtered.groupby("misspelling_type").agg(counts=("normalized_query","count"), avg_ctr=("CTR","mean")).reset_index()
        st.plotly_chart(px.bar(m.sort_values("counts", ascending=False), x="misspelling_type", y="counts", title="Misspelling Type Counts"), use_container_width=True)
        st.plotly_chart(px.bar(m.sort_values("avg_ctr", ascending=False), x="misspelling_type", y="avg_ctr", title="Avg CTR by Misspelling Type"), use_container_width=True)

    # 13 Sub Category revenue share
    if "Sub Category" in df_filtered.columns:
        st.subheader("13) Revenue by Sub Category (Top 15)")
        st.plotly_chart(px.bar(df_filtered.groupby("Sub Category")["Rev"].sum().reset_index().sort_values("Rev", ascending=False).head(15), x="Sub Category", y="Rev", title="Top Sub Categories by Revenue"), use_container_width=True)

    # 14 Brands from product master (if present)
    if "brand" in df_filtered.columns or "brand_canonical" in df_filtered.columns:
        brand_col = "brand_canonical" if "brand_canonical" in df_filtered.columns else "brand"
        st.subheader("14) Top Brands by Query Count")
        st.plotly_chart(px.bar(df_filtered.groupby(brand_col)["Count"].sum().reset_index().sort_values("Count", ascending=False).head(20), x=brand_col, y="Count", title="Top Brands by Search Count"), use_container_width=True)

# ---------- Temporal Tab ----------
with tabs[4]:
    st.header("Temporal & Monthly analysis")
    # 15 CTR trend by Month
    st.subheader("15) Average CTR by Month")
    ct = df_filtered.groupby("Month_Year")["CTR"].mean().reset_index()
    st.plotly_chart(px.line(ct.sort_values("Month_Year"), x="Month_Year", y="CTR", markers=True, title="Average CTR by Month"), use_container_width=True)

    # 16 Revenue trend by Month
    st.subheader("16) Total Revenue by Month")
    rv = df_filtered.groupby("Month_Year")["Rev"].sum().reset_index()
    st.plotly_chart(px.line(rv.sort_values("Month_Year"), x="Month_Year", y="Rev", markers=True, title="Revenue by Month"), use_container_width=True)

    # 17 Top rising queries MoM (pct change)
    st.subheader("17) Top rising queries month-over-month (by impressions)")
    mom = df_filtered.groupby(["Month_Year","normalized_query"])["Count"].sum().reset_index()
    mom_pivot = mom.pivot(index="normalized_query", columns="Month_Year", values="Count").fillna(0)
    if mom_pivot.shape[1] >= 2:
        months_sorted = sorted(mom_pivot.columns)
        recent, prev = months_sorted[-1], months_sorted[-2]
        mom_pivot["pct_change"] = (mom_pivot[recent] - mom_pivot[prev]) / (mom_pivot[prev].replace({0:np.nan}))
        rising = mom_pivot.sort_values("pct_change", ascending=False).reset_index().head(15)[["normalized_query","pct_change"]]
        st.dataframe(rising.rename(columns={"normalized_query":"query","pct_change":"MoM_pct_change"}), use_container_width=True)
    else:
        st.info("Not enough months to compute MoM.")

    # 18 Weekday revenue
    st.subheader("18) Revenue by Day of Week")
    if "day_of_week" in df_filtered.columns:
        dow = df_filtered.groupby("day_of_week")["Rev"].sum().reset_index()
        # ensure ordering
        order = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
        dow["day_of_week"] = pd.Categorical(dow["day_of_week"], categories=order, ordered=True)
        st.plotly_chart(px.bar(dow.sort_values("day_of_week"), x="day_of_week", y="Rev", title="Revenue by Day of Week"), use_container_width=True)

# ---------- Insights Hub (30 analyses) ----------
with tabs[5]:
    st.header("💡 Insights Hub — 30 Impactful Questions (focus on search column)")

    insights = [
        ("Q1", "What are the top 20 search queries by impressions?", "Shows what customers search most — prioritize inventory & search tuning."),
        ("Q2", "Which queries have high impressions but low CTR?", "Opportunity: improve search results or meta content to increase CTR."),
        ("Q3", "Which queries have high CTR but low revenue?", "High interest but low purchase: check pricing, availability, or landing pages."),
        ("Q4", "What are the top revenue-driving queries?", "Focus promotions & stock for these queries."),
        ("Q5", "Which queries convert best (high CR)?", "High intent queries — prioritize for paid campaigns."),
        ("Q6", "What words appear most in searches?", "Inform SEO, synonyms, synonyms to add to product titles/descriptions."),
        ("Q7", "How does query length affect CTR/CR?", "Short vs long query behavior indicates how users search."),
        ("Q8", "Which misspelling types hurt CTR most?", "Improve fuzzy search / suggestions for those misspelling types."),
        ("Q9", "Are there specific brands dominating search terms?", "Measure brand demand and stock/marketing alignment."),
        ("Q10", "Top rising queries month-over-month?", "Spot emerging demand and seasonal opportunities."),
        ("Q11", "Which queries are seasonal across months?", "Plan promotions and inventory ahead of peaks."),
        ("Q12", "Which queries have a high click-to-conversion drop (high CTR, low CR)?", "Investigate landing pages or product availability."),
        ("Q13", "Which queries have high conversion value per conversion?", "High-LTV queries to push in ads."),
        ("Q14", "Which queries are unique to Arabic vs other languages?", "Localize content and landing pages."),
        ("Q15", "What is the revenue share of top 10 queries?", "Concentration of revenue among few queries."),
        ("Q16", "Which queries are frequent but return low revenue per search?", "Optimization candidates to drive revenue uplift."),
        ("Q17", "Which subcategories produce the best conversion rates?", "Category prioritization for merchandising."),
        ("Q18", "Which queries show anomalous spikes in a month?", "Detect marketing or data issues."),
        ("Q19", "Which queries have low match confidence (bad matches)?", "Improve matching rules or product feed."),
        ("Q20", "How many unique queries are there per month?", "Measure search breadth and new query growth."),
        ("Q21", "Which queries have the largest drop in CTR MoM?", "Investigate ranking or product availability changes."),
        ("Q22", "What queries show the best revenue per impression?", "Most efficient queries from organic traffic."),
        ("Q23", "Which queries cause many zero-conversion clicks?", "Check funnel & product detail page."),
        ("Q24", "Which queries show repeated misspellings but similar intent?", "Teach synonyms & synonyms mapping."),
        ("Q25", "Which queries are high volume on weekends vs weekdays?", "Weekend-specific campaigns."),
        ("Q26", "Which queries have low visibility but high conversion rate?", "Hidden gems to surface more."),
        ("Q27", "Which brands have the most search share growth MoM?", "Brand momentum tracking."),
        ("Q28", "Which queries drive most of the returns/cancellations (if field exists)?", "Product quality / expectation mismatch signals."),
        ("Q29", "Which queries have the highest ATCR (Add-to-cart conversion rate)?", "High intent to buy — optimize checkout."),
        ("Q30", "Which queries are long-tail but collectively high volume?", "Long-tail SEO and content opportunities.")
    ]

    # Render insights with small visual sample and suggested action
    for qid, qtext, qinsight in insights:
        with st.expander(f"{qid}: {qtext}"):
            st.write(qinsight)
            # show a small sample visualization related to the question
            if qid == "Q1":
                st.plotly_chart(px.bar(top_n_queries(df_filtered,20), x="query", y="count", title=qtext), use_container_width=True)
            elif qid == "Q2":
                tmp = df_filtered.copy()
                tmp["impr_rank"] = tmp["Count"].rank(ascending=False, method="first")
                subset = tmp[(tmp["Count"]>=tmp["Count"].quantile(0.6)) & (tmp["CTR"]<=tmp["CTR"].quantile(0.3))]
                st.dataframe(subset.sort_values(["Count","CTR"], ascending=[False,True])[["normalized_query","Count","CTR","Rev"]].head(20))
            elif qid == "Q10":
                # show top rising MoM (from earlier)
                if 'pct_change' in locals():
                    st.dataframe(rising.rename(columns={"normalized_query":"query","pct_change":"MoM_pct_change"}), use_container_width=True)
                else:
                    st.info("Not enough months to compute MoM.")
            elif qid == "Q6":
                st.plotly_chart(px.bar(most_common_words(df_sample, top_n=30).head(15), x="word", y="count", title="Top words in queries"), use_container_width=True)
            else:
                # default table showing top related queries
                st.dataframe(df_filtered.groupby("normalized_query").agg(impressions=("Count","sum"), avg_CTR=("CTR","mean"), revenue=("Rev","sum")).reset_index().sort_values("impressions", ascending=False).head(15), use_container_width=True)

# ---------- Pivot & Export ----------
with tabs[6]:
    st.header("📋 Pivot Tables & Export")
    st.markdown("Interactive pivots — select indices/columns and aggregation to generate cross-tabs. Use AgGrid for best experience (enable in sidebar).")

    # Pivot builder UI
    all_cols = df_filtered.columns.tolist()
    idx = st.multiselect("Pivot index (rows)", options=all_cols, default=["normalized_query"])
    cols = st.multiselect("Pivot columns", options=all_cols, default=["language"])
    val = st.selectbox("Value (measure)", options=["Count","Rev","CTR","CR","conversions","conversion_value","revenue_per_search"], index=0)
    agg = st.selectbox("Aggregation", options=["sum","mean","count"], index=0)
    if st.button("Generate pivot"):
        pivot_df = build_pivot(df_filtered, index=idx, columns=cols, values=val, aggfunc=agg)
        if pivot_df.empty:
            st.warning("Pivot returned empty or failed. Try fewer columns/indices.")
        else:
            st.success("Pivot generated — you can download or inspect.")
            if AGGRID_OK and show_aggrid:
                aggrid_display(pivot_df, height=500)
            else:
                st.dataframe(pivot_df, use_container_width=True)
            csv = pivot_df.to_csv(index=False).encode('utf-8')
            st.download_button("⬇ Download pivot CSV", csv, file_name="pivot_export.csv")

    st.markdown("---")
    st.subheader("Download filtered dataset")
    st.download_button("⬇ Download filtered CSV", df_filtered.to_csv(index=False).encode('utf-8'), file_name="filtered_search_data.csv")

# ---------- Footer / Tips ----------
st.markdown("---")
st.markdown("""
#### Tips
- Use the **Insights Hub** to get quick business questions and starter actions.
- Use **Pivot & Export** to create ad-hoc tables for stakeholders.
- If AgGrid or WordCloud isn't available, enable them in your environment (`pip install streamlit-aggrid wordcloud`) for better interactivity.
""")
