import io
import re
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import requests
from bs4 import BeautifulSoup

# ── Header ───────────────────────────────────────────────
st.title("Multi-Page Web Data Extractor")
st.write("Extract data from multiple pages and download a clean Excel report.")

# ── Sidebar ──────────────────────────────────────────────
st.sidebar.header("Settings")

mode = st.sidebar.selectbox(
    "Mode",
    ["Extract Data Only", "Full Business Report"]
)

scraping_method = st.sidebar.radio(
    "Scraping Mode",
    ["Auto (Tables)", "Smart Extract (Books / Products)", "Advanced (CSS Selector)"]
)

total_pages = st.sidebar.number_input(
    "Number of pages",
    min_value=1,
    max_value=50,
    value=3
)

url_template = st.sidebar.text_input(
    "URL pattern (use {page})",
    placeholder="https://example.com/page-{page}.html"
)

if not url_template:
    st.info("Enter a URL pattern to start.")
    st.stop()

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# ── Validate first page ──────────────────────────────────
first_url = url_template.replace("{page}", "1")

try:
    first_response = requests.get(first_url, headers=headers, timeout=10)
    html = first_response.text
except Exception:
    st.error("❌ Could not access the website.")
    st.stop()

if first_response.status_code != 200:
    st.error(f"❌ Access denied ({first_response.status_code})")
    st.stop()

# ── Table selection (Auto) ───────────────────────────────
selected_table_index = 0

if scraping_method == "Auto (Tables)":
    try:
        tables = pd.read_html(io.StringIO(html))
        st.success(f"✅ Found {len(tables)} table(s)")
        selected_table_index = st.sidebar.selectbox(
            "Select table",
            range(len(tables)),
            format_func=lambda x: f"Table {x+1}"
        )
    except Exception:
        st.warning("⚠️ No tables found. Try Smart Extract or Advanced mode.")
        st.stop()

# ── Advanced mode input ──────────────────────────────────
selector = None
if scraping_method == "Advanced (CSS Selector)":
    selector = st.text_input("CSS Selector (e.g. div.title)")
    if not selector:
        st.info("Enter a CSS selector.")
        st.stop()

# ── SCRAPE ALL PAGES ─────────────────────────────────────
all_data = []
progress = st.progress(0)
status = st.empty()

rating_map = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

for i in range(1, total_pages + 1):
    url = url_template.replace("{page}", str(i))
    status.text(f"Scraping page {i}/{total_pages}...")

    try:
        res = requests.get(url, headers=headers, timeout=10)
        page_html = res.text

        if scraping_method == "Auto (Tables)":
            tables = pd.read_html(io.StringIO(page_html))
            all_data.append(tables[selected_table_index])

        elif scraping_method == "Smart Extract (Books / Products)":
            soup = BeautifulSoup(page_html, "html.parser")
            products = soup.select("article.product_pod")

            if products:
                rows = []
                for p in products:
                    title_tag = p.select_one("h3 a")
                    title = title_tag["title"] if title_tag and title_tag.has_attr("title") else (title_tag.get_text(strip=True) if title_tag else "N/A")

                    price_tag = p.select_one("p.price_color")
                    price_text = price_tag.get_text(strip=True) if price_tag else "0"
                    price = float(re.sub(r"[^\d.]", "", price_text)) if price_text else 0.0

                    rating_tag = p.select_one("p.star-rating")
                    rating_word = rating_tag["class"][1] if rating_tag and len(rating_tag["class"]) > 1 else "Zero"
                    rating = rating_map.get(rating_word, 0)

                    avail_tag = p.select_one("p.availability")
                    availability = avail_tag.get_text(strip=True) if avail_tag else "N/A"

                    rows.append({
                        "Title": title,
                        "Price (£)": price,
                        "Rating": rating,
                        "Availability": availability
                    })

                all_data.append(pd.DataFrame(rows))
            else:
                st.warning(f"⚠️ No products found on page {i}")
                continue

        else:
            soup = BeautifulSoup(page_html, "html.parser")
            elements = soup.select(selector)
            if not elements:
                st.warning(f"No data on page {i}")
                continue
            data = [el.get_text(strip=True) for el in elements]
            all_data.append(pd.DataFrame(data, columns=["Extracted Data"]))

    except Exception:
        st.warning(f"⚠️ Failed page {i}")
        continue

    progress.progress(int((i / total_pages) * 50))

status.empty()

# ── Combine ──────────────────────────────────────────────
if not all_data:
    st.error("❌ No data scraped.")
    st.stop()

df_raw = pd.concat(all_data, ignore_index=True)

st.subheader("🔍 Raw Preview")
st.dataframe(df_raw.head(20), hide_index=True)

# ── PROCESS BUTTON ───────────────────────────────────────
if st.button("🚀 Generate Report"):

    prog = st.progress(0)

    with st.spinner("Processing..."):

        df = df_raw.copy()
        prog.progress(20)

        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        empty_cells = int(df.isnull().sum().sum())

        numeric_cols = df.select_dtypes(include="number").columns
        df[numeric_cols] = df[numeric_cols].fillna(0)
        df = df.fillna("N/A")

        prog.progress(40)

        sort_col = st.sidebar.selectbox("Sort by", ["None"] + list(df.columns), key="sort_col")
        if sort_col != "None":
            order = st.sidebar.selectbox("Order", ["Descending", "Ascending"], key="sort_order")
            try:
                df = df.sort_values(sort_col, ascending=(order == "Ascending"))
            except Exception:
                st.warning("⚠️ Sorting failed — mixed data types.")

        prog.progress(60)

        st.subheader("📊 Clean Data")
        st.dataframe(df, hide_index=True)

        summary_df = None

        if mode == "Full Business Report":
            num_cols = list(df.select_dtypes(include="number").columns)
            cat_cols = list(df.select_dtypes(exclude="number").columns)

            if not num_cols or not cat_cols:
                st.warning("⚠️ Not enough data for summary. Need at least one text and one numeric column.")
            else:
                group_col = st.sidebar.selectbox("Group by", cat_cols, key="group_col")
                value_col = st.sidebar.selectbox("Analyze", num_cols, key="value_col")

                df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

                summary_df = df.groupby(group_col)[value_col].agg(
                    Total="sum",
                    Average="mean",
                    Highest="max",
                    Count="count"
                ).reset_index()

                summary_df["Total"] = summary_df["Total"].round(2)
                summary_df["Average"] = summary_df["Average"].round(2)
                summary_df["Highest"] = summary_df["Highest"].round(2)

                summary_sort = st.sidebar.selectbox(
                    "Sort summary by",
                    ["None", "Total", "Average", "Highest", "Count"],
                    key="summary_sort"
                )
                if summary_sort != "None":
                    summary_order = st.sidebar.selectbox(
                        "Summary order",
                        ["Descending", "Ascending"],
                        key="summary_order"
                    )
                    summary_df = summary_df.sort_values(
                        summary_sort,
                        ascending=(summary_order == "Ascending")
                    )

                st.subheader("📈 Summary Report")
                st.dataframe(summary_df, hide_index=True)

                chart_metric = st.selectbox(
                    "Chart metric",
                    ["Total", "Average", "Highest", "Count"],
                    key="chart_metric"
                )
                chart_data = summary_df.set_index(group_col)[[chart_metric]]
                st.subheader(f"{chart_metric} by {group_col}")
                st.bar_chart(chart_data)

        prog.progress(80)

        # ── Excel ─────────────────────────────────────────
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            df.to_excel(writer, index=False, sheet_name="Scraped Data")
            if summary_df is not None:
                summary_df.to_excel(writer, index=False, sheet_name="Summary")

        buffer.seek(0)
        wb = load_workbook(buffer)

        for sheet in wb.sheetnames:
            for cell in wb[sheet][1]:
                cell.font = Font(bold=True)

        final = io.BytesIO()
        wb.save(final)
        final.seek(0)
        prog.progress(100)

    col1, col2, col3 = st.columns(3)
    col1.metric("Rows Extracted", len(df))
    col2.metric("Pages Scraped", total_pages)
    col3.metric("Empty Cells Fixed", empty_cells)

    st.success("✅ Report ready!")
    st.download_button("⬇️ Download Excel", final.getvalue(), "scraped_data.xlsx")
