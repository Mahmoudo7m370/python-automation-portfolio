import io
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import requests
from bs4 import BeautifulSoup

# ── Page Header ───────────────────────────────────────────────
st.title("Multi-Page Web Scraper Pro")
st.write("Scrape multiple pages from any website and get a clean, formatted Excel report.")

# ── Sidebar: Settings ─────────────────────────────────────────
st.sidebar.header("Settings")

report_mode = st.sidebar.selectbox(
    "Report mode",
    ["Scraping Only", "Scraping + Summary Report"]
)

scraping_method = st.sidebar.radio(
    "Scraping method",
    ["Auto — detect tables", "Manual — enter tag and class"]
)

total_pages = st.sidebar.number_input(
    "Number of pages to scrape",
    min_value=1,
    max_value=50,
    value=1
)

url_template = st.sidebar.text_input(
    "URL pattern — use {page} where the page number goes",
    placeholder="https://example.com/page-{page}.html"
)

# ── Wait for URL ──────────────────────────────────────────────
if not url_template:
    st.info("Enter a URL pattern in the sidebar to get started.")
    st.info("Example: https://books.toscrape.com/catalogue/page-{page}.html")
    st.stop()

# ── Validate First Page ───────────────────────────────────────
first_page_url = url_template.replace("{page}", "1")

try:
    headers = {"User-Agent": "Mozilla/5.0"}
    first_page_response = requests.get(first_page_url, timeout=10,headers=headers)
except Exception:
    st.error("Could not reach this URL. Check the address and try again.")
    st.stop()

if first_page_response.status_code != 200:
    st.error(f"Access denied — server returned status {first_page_response.status_code}.")
    st.stop()

# ── Determine Table Selection (Auto mode only) ────────────────
selected_table_index = 0

if scraping_method == "Auto — detect tables":
    try:
        first_page_tables = pd.read_html(io.StringIO(first_page_response.text))
        if len(first_page_tables) > 1:
            selected_table_index = st.sidebar.selectbox(
                f"Found {len(first_page_tables)} tables — which one?",
                range(len(first_page_tables)),
                format_func=lambda x: f"Table {x + 1}"
            )
    except Exception:
        st.error("No tables found on page 1. Switch to Manual mode.")
        st.stop()

# ── Manual Mode Inputs ────────────────────────────────────────
html_tag = None
css_class = None

if scraping_method == "Manual — enter tag and class":
    html_tag = st.text_input("HTML tag (e.g. div, span, h2)")
    css_class = st.text_input("CSS class (e.g. product-title)")

    if not html_tag or not css_class:
        st.info("Enter the HTML tag and CSS class to start scraping.")
        st.stop()

# ── Scrape All Pages ──────────────────────────────────────────
all_page_dataframes = []
scraping_progress = st.progress(0)
scraping_status = st.empty()

for page_number in range(1, total_pages + 1):
    current_url = url_template.replace("{page}", str(page_number))
    scraping_status.text(f"Scraping page {page_number} of {total_pages}...")
    scraping_progress.progress(int((page_number / total_pages) * 50))

    if scraping_method == "Auto — detect tables":
        try:
            page_response = requests.get(current_url, timeout=10, headers=headers)
            page_tables =pd.read_html(io.StringIO(page_response.text))
            all_page_dataframes.append(page_tables[selected_table_index])
        except Exception:
            st.warning(f"No tables found on page {page_number} — skipping.")
            continue

    elif scraping_method == "Manual — enter tag and class":
        try:
            page_response = requests.get(current_url, timeout=10,headers=headers)
            page_soup = BeautifulSoup(page_response.text, "html.parser")
            matched_elements = page_soup.find_all(html_tag, class_=css_class)

            if not matched_elements:
                st.warning(f"No elements found on page {page_number} — skipping.")
                continue

            extracted_text = [element.text.strip() for element in matched_elements]
            page_dataframe = pd.DataFrame(extracted_text, columns=["Scraped Data"])
            all_page_dataframes.append(page_dataframe)
        except Exception:
            st.warning(f"Could not scrape page {page_number} — skipping.")
            continue

scraping_status.empty()

# ── Combine All Pages ─────────────────────────────────────────
if not all_page_dataframes:
    st.error("No data was scraped. Check your URL pattern and settings.")
    st.stop()

combined_df = pd.concat(all_page_dataframes, ignore_index=True)
st.success(f"Scraped {len(combined_df)} rows across {total_pages} page(s).")

# ── Process and Report ────────────────────────────────────────
enable_highlight = False
processing_progress = st.progress(0)

with st.spinner("Processing your data..."):

    # Step 1 — Work on a copy of the combined data
    working_df = combined_df.copy()
    processing_progress.progress(15)

    # Step 2 — Count empty cells before filling
    empty_cell_count = working_df.isnull().sum().sum()

    # Step 3 — Clean text columns: strip spaces and fix casing
    for column in working_df.columns:
        try:
            working_df[column] = working_df[column].str.strip().str.title()
        except Exception:
            pass

    # Step 4 — Validate summary mode requirements
    numeric_columns = working_df.select_dtypes(include="number").columns
    has_numeric_columns = len(numeric_columns) > 0

    if not has_numeric_columns and report_mode == "Scraping + Summary Report":
        st.error("Summary mode requires at least one numeric column. This data has none.")
        st.stop()

    # Step 5 — Fill empty cells
    working_df[numeric_columns] = working_df[numeric_columns].fillna(0)
    working_df = working_df.fillna("N/A")
    processing_progress.progress(40)

    # Step 6 — Optional sorting
    sort_by_column = st.sidebar.selectbox(
        "Sort by column",
        ["None"] + list(working_df.columns),
        key="sort_column"
    )
    if sort_by_column != "None":
        sort_direction = st.sidebar.selectbox(
            "Sort order",
            ["Ascending", "Descending"],
            key="sort_direction"
        )
        working_df = working_df.sort_values(
            sort_by_column,
            ascending=(sort_direction == "Ascending")
        )
    processing_progress.progress(60)

    # Step 7 — Display scraped data
    st.subheader("Scraped Data")
    st.dataframe(working_df)

    summary_df = None

    # Step 8 — Build summary report if selected
    if report_mode == "Scraping + Summary Report":

        amount_column = st.sidebar.selectbox(
            "Column to summarize",
            working_df.select_dtypes(include="number").columns,
            key="summary_column"
        )
        working_df[amount_column] = pd.to_numeric(working_df[amount_column], errors="coerce")
        groupby_column = working_df.select_dtypes(exclude="number").columns[0]

        summary_df = working_df.groupby(groupby_column)[amount_column].agg(
            Total="sum",
            Average="mean",
            Highest="max"
        ).reset_index()
        summary_df["Average"] = summary_df["Average"].round()

        summary_sort_column = st.sidebar.selectbox(
            "Sort summary by",
            ["None", "Total", "Average", "Highest"],
            key="summary_sort_column"
        )
        if summary_sort_column != "None":
            summary_sort_direction = st.sidebar.selectbox(
                "Summary sort order",
                ["Ascending", "Descending"],
                key="summary_sort_direction"
            )
            summary_df = summary_df.sort_values(
                summary_sort_column,
                ascending=(summary_sort_direction == "Ascending")
            )

        st.subheader("Summary Report")
        st.dataframe(summary_df)

    # Step 9 — Optional row highlight
    enable_highlight = st.sidebar.checkbox("Highlight a row")
    if enable_highlight:
        highlight_row_number = st.sidebar.number_input(
            "Row number to highlight (1 = first row)",
            min_value=1,
            max_value=len(working_df),
            step=1,
            key="highlight_row"
        )
        highlight_color = st.sidebar.color_picker(
            "Highlight color",
            "#FFFF00",
            key="highlight_color"
        )

    processing_progress.progress(80)

    # Step 10 — Write data to Excel in memory
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer) as excel_writer:
        working_df.to_excel(excel_writer, sheet_name="Scraped Data", index=False)
        if summary_df is not None:
            summary_df.to_excel(excel_writer, sheet_name="Summary", index=False)

    # Step 11 — Apply formatting with openpyxl
    excel_buffer.seek(0)
    workbook = load_workbook(excel_buffer)
    data_sheet = workbook["Scraped Data"]

    for cell in data_sheet[1]:
        cell.font = Font(bold=True)

    if summary_df is not None:
        for cell in workbook["Summary"][1]:
            cell.font = Font(bold=True)

    if enable_highlight:
        fill_color = highlight_color[1:]
        row_fill = PatternFill(
            start_color=fill_color,
            end_color=fill_color,
            fill_type="solid"
        )
        for cell in data_sheet[int(highlight_row_number) + 1]:
            cell.fill = row_fill

    # Step 12 — Save final workbook to memory
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    processing_progress.progress(100)

# ── Results ───────────────────────────────────────────────────
st.metric("Total Rows Scraped", len(working_df))
st.metric("Pages Scraped", total_pages)
st.metric("Empty Cells Filled", empty_cell_count)
st.success("Your report is ready to download!")
st.download_button(
    "Download Excel Report",
    final_output.getvalue(),
    "scraped_report.xlsx"
)