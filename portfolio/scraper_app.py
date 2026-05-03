import io
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import requests
from bs4 import BeautifulSoup

# ── Page Header ───────────────────────────────────────────────
st.title("Web Data Extractor")
st.write("Extract data from any website and download a clean Excel report.")

# ── Sidebar ───────────────────────────────────────────────────
st.sidebar.header("Settings")

mode = st.sidebar.selectbox(
    "Select Mode",
    ["Extract Data Only", "Full Business Report"]
)

scraping_method = st.sidebar.radio(
    "Scraping Mode",
    ["Auto (Tables)", "Advanced (CSS Selector)"]
)

target_url = st.sidebar.text_input("Enter Website URL")

# ── Stop if no URL ────────────────────────────────────────────
if not target_url:
    st.info("Enter a URL to begin.")
    st.stop()

# ── Fetch Page ────────────────────────────────────────────────
try:
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(target_url, timeout=10, headers=headers)
    html = response.text
except Exception:
    st.error("❌ Could not access this URL.")
    st.stop()

if response.status_code != 200:
    st.error(f"❌ Access denied (Status {response.status_code})")
    st.stop()

# ── Scraping ──────────────────────────────────────────────────
scraped_df = None

if scraping_method == "Auto (Tables)":
    try:
        tables = pd.read_html(io.StringIO(html))
        st.success(f"✅ Found {len(tables)} table(s)")

        table_index = st.sidebar.selectbox(
            "Select Table",
            range(len(tables)),
            format_func=lambda x: f"Table {x+1}"
        )

        scraped_df = tables[table_index]

    except Exception:
        st.warning("⚠️ No tables detected. Try Advanced mode.")

elif scraping_method == "Advanced (CSS Selector)":
    selector = st.text_input("Enter CSS Selector (e.g. div.product-title)")

    if selector:
        soup = BeautifulSoup(html, "html.parser")
        elements = soup.select(selector)

        if not elements:
            st.error("❌ No elements found.")
            st.code(html[:1000])  # show snippet for debugging
            st.stop()

        data = [el.get_text(strip=True) for el in elements]
        scraped_df = pd.DataFrame(data, columns=["Extracted Data"])

    else:
        st.info("Enter a CSS selector to continue.")
        st.stop()

# ── Preview BEFORE processing ─────────────────────────────────
if scraped_df is not None:
    st.subheader("🔍 Preview (Raw Data)")
    st.dataframe(scraped_df.head(20))

# ── Processing ────────────────────────────────────────────────
if scraped_df is not None and len(scraped_df) > 0:

    if st.button("🚀 Generate Report"):

        progress = st.progress(0)

        with st.spinner("Processing..."):

            df = scraped_df.copy()
            progress.progress(20)

            # Clean text
            for col in df.columns:
                if df[col].dtype == "object":
                    df[col] = df[col].astype(str).str.strip()

            empty_cells = df.isnull().sum().sum()

            # Fill missing
            numeric_cols = df.select_dtypes(include="number").columns
            df[numeric_cols] = df[numeric_cols].fillna(0)
            df = df.fillna("N/A")

            progress.progress(40)

            # Sorting (SAFE)
            sort_col = st.sidebar.selectbox(
                "Sort by",
                ["None"] + list(df.columns)
            )

            if sort_col != "None":
                order = st.sidebar.selectbox("Order", ["Ascending", "Descending"])
                try:
                    df = df.sort_values(sort_col, ascending=(order == "Ascending"))
                except Exception:
                    st.warning("⚠️ Could not sort due to mixed data types.")

            progress.progress(60)

            st.subheader("📊 Clean Data")
            st.dataframe(df)

            summary_df = None

            # ── Summary Report ───────────────────────────────
            if mode == "Full Business Report":

                numeric_cols = df.select_dtypes(include="number").columns
                non_numeric_cols = df.select_dtypes(exclude="number").columns

                if len(numeric_cols) == 0 or len(non_numeric_cols) == 0:
                    st.warning("⚠️ Not enough data for summary.")
                else:
                    group_col = st.sidebar.selectbox("Group by", non_numeric_cols)
                    value_col = st.sidebar.selectbox("Analyze column", numeric_cols)

                    df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

                    summary_df = df.groupby(group_col)[value_col].agg(
                        Total="sum",
                        Average="mean",
                        Max="max"
                    ).reset_index()

                    st.subheader("📈 Summary")
                    st.dataframe(summary_df)

            progress.progress(80)

            # ── Excel Export ────────────────────────────────
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                df.to_excel(writer, index=False, sheet_name="Data")
                if summary_df is not None:
                    summary_df.to_excel(writer, index=False, sheet_name="Summary")

            buffer.seek(0)
            wb = load_workbook(buffer)

            # Bold headers
            for sheet in wb.sheetnames:
                for cell in wb[sheet][1]:
                    cell.font = Font(bold=True)

            final = io.BytesIO()
            wb.save(final)
            final.seek(0)

            progress.progress(100)

        # ── Output ────────────────────────────────────────
        st.metric("Rows", len(df))
        st.metric("Empty Cells Fixed", empty_cells)

        st.success("✅ Your report is ready!")

        st.download_button(
            "⬇️ Download Excel",
            final.getvalue(),
            "report.xlsx"
        )

else:
    st.info("No data extracted yet.")