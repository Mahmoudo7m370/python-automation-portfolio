import io
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
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
    ["Auto (Tables)", "Advanced (CSS Selector)"]
)

total_pages = st.sidebar.number_input(
    "Number of pages",
    min_value=1,
    max_value=50,
    value=1
)

url_template = st.sidebar.text_input(
    "URL pattern (use {page})",
    placeholder="https://example.com/page-{page}.html"
)

if not url_template:
    st.info("Enter a URL pattern to start.")
    st.stop()

headers = {"User-Agent": "Mozilla/5.0"}

# ── Validate first page ──────────────────────────────────
first_url = url_template.replace("{page}", "1")

try:
    first_response = requests.get(first_url, headers=headers, timeout=10)
    html = first_response.text
except:
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
        st.success(f"Found {len(tables)} table(s)")

        selected_table_index = st.sidebar.selectbox(
            "Select table",
            range(len(tables)),
            format_func=lambda x: f"Table {x+1}"
        )

    except:
        st.warning("No tables found. Try Advanced mode.")
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

for i in range(1, total_pages + 1):
    url = url_template.replace("{page}", str(i))
    status.text(f"Scraping page {i}/{total_pages}...")

    try:
        res = requests.get(url, headers=headers, timeout=10)
        page_html = res.text

        if scraping_method == "Auto (Tables)":
            tables = pd.read_html(io.StringIO(page_html))
            all_data.append(tables[selected_table_index])

        else:
            soup = BeautifulSoup(page_html, "html.parser")
            elements = soup.select(selector)

            if not elements:
                st.warning(f"No data on page {i}")
                continue

            data = [el.get_text(strip=True) for el in elements]
            all_data.append(pd.DataFrame(data, columns=["Extracted Data"]))

    except Exception:
        st.warning(f"Failed page {i}")
        continue

    progress.progress(int((i / total_pages) * 50))

status.empty()

# ── Combine ─────────────────────────────────────────────
if not all_data:
    st.error("No data scraped.")
    st.stop()

df_raw = pd.concat(all_data, ignore_index=True)

# ── PREVIEW FIRST (VERY IMPORTANT) ──────────────────────
st.subheader("🔍 Raw Preview")
st.dataframe(df_raw.head(20))

# ── PROCESS BUTTON ─────────────────────────────────────
if st.button("🚀 Generate Report"):

    prog = st.progress(0)

    with st.spinner("Processing..."):

        df = df_raw.copy()
        prog.progress(20)

        # Clean text
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        empty_cells = df.isnull().sum().sum()

        # Fill missing
        numeric_cols = df.select_dtypes(include="number").columns
        df[numeric_cols] = df[numeric_cols].fillna(0)
        df = df.fillna("N/A")

        prog.progress(40)

        # SAFE SORTING
        sort_col = st.sidebar.selectbox("Sort by", ["None"] + list(df.columns))
        if sort_col != "None":
            order = st.sidebar.selectbox("Order", ["Ascending", "Descending"])
            try:
                df = df.sort_values(sort_col, ascending=(order == "Ascending"))
            except:
                st.warning("Sorting failed due to mixed data types.")

        prog.progress(60)

        st.subheader("📊 Clean Data")
        st.dataframe(df)

        summary_df = None

        # ── SUMMARY (FIXED) ─────────────────────────────
        if mode == "Full Business Report":

            num_cols = df.select_dtypes(include="number").columns
            cat_cols = df.select_dtypes(exclude="number").columns

            if len(num_cols) == 0 or len(cat_cols) == 0:
                st.warning("Not enough data for summary.")
            else:
                group_col = st.sidebar.selectbox("Group by", cat_cols)
                value_col = st.sidebar.selectbox("Analyze", num_cols)

                df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

                summary_df = df.groupby(group_col)[value_col].agg(
                    Total="sum",
                    Average="mean",
                    Max="max"
                ).reset_index()

                st.subheader("📈 Summary")
                st.dataframe(summary_df)

        prog.progress(80)

        # ── Excel ──────────────────────────────────────
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
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

    # ── Output ───────────────────────────────────────
    st.metric("Rows", len(df))
    st.metric("Pages Scraped", total_pages)
    st.metric("Empty Fixed", empty_cells)

    st.success("Report ready!")

    st.download_button(
        "⬇️ Download Excel",
        final.getvalue(),
        "report.xlsx"
    )