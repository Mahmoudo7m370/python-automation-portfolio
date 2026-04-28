import io
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# ── Page Header ───────────────────────────────────────────────
st.title("Data Cleaner Pro")
st.write("Upload your messy CSV or Excel files and get a clean, formatted Excel report.")

# ── User Selects Mode ──────────────────────────────────────────
selected_mode = st.selectbox("Choose mode", ["Clean Only", "Clean + Summary Report"])

# ── File Upload ────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload CSV or Excel files",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    # ── Read and Combine All Uploaded Files ────────────────────
    dataframes = []
    for uploaded_file in uploaded_files:
        try:
            dataframes.append(pd.read_csv(uploaded_file))
        except Exception:
            dataframes.append(pd.read_excel(uploaded_file))

    combined_df = pd.concat(dataframes, ignore_index=True)

    # ── Track Original Row Count and Empty Cells ───────────────
    original_row_count = len(combined_df)
    empty_cell_count = combined_df.isnull().sum().sum()

    # ── Clean Text Columns ─────────────────────────────────────
    for column in combined_df.columns:
        try:
            combined_df[column] = combined_df[column].str.strip().str.title()
        except Exception:
            pass

    # ── Remove Duplicate Rows ──────────────────────────────────
    combined_df = combined_df.drop_duplicates()

    # ── Fill Empty Cells ───────────────────────────────────────
    numeric_columns = combined_df.select_dtypes(include="number").columns
    combined_df[numeric_columns] = combined_df[numeric_columns].fillna(0)
    combined_df = combined_df.fillna("N/A")

    # ── Optional: Sort the Data ────────────────────────────────
    sort_data = st.selectbox(
        "Do you want to sort the data?",
        ["no", "yes"],
        key="main_sort_choice"
    )
    if sort_data == "yes":
        sort_column = st.selectbox("Choose column to sort", combined_df.columns, key="main_sort_column")
        sort_direction = st.selectbox("Order", ["ascending", "descending"], key="main_sort_direction")
        combined_df = combined_df.sort_values(
            sort_column,
            ascending=(sort_direction == "ascending")
        )

    # ── Display Cleaned Data ───────────────────────────────────
    st.subheader("Cleaned Data")
    st.dataframe(combined_df)

    # ── Calculate Duplicates Removed ──────────────────────────
    duplicates_removed = original_row_count - len(combined_df)
    summary_df = None

    # ── Summary Report Mode ────────────────────────────────────
    if selected_mode == "Clean + Summary Report":

        # User picks which numeric column to summarize
        amount_column = st.selectbox(
            "Choose the numeric column to summarize",
            combined_df.select_dtypes(include="number").columns,
            key="summary_col"
        )

        # Convert chosen column to numeric safely
        combined_df[amount_column] = pd.to_numeric(combined_df[amount_column], errors="coerce")

        # Automatically detect the first text column to group by
        groupby_column = combined_df.select_dtypes(exclude="number").columns[0]

        # Build summary table
        summary_df = combined_df.groupby(groupby_column)[amount_column].agg(
            Total="sum",
            Average="mean",
            Highest="max"
        ).reset_index()
        summary_df["Average"] = summary_df["Average"].round()

        # Optional: Sort the summary
        sort_summary = st.selectbox(
            "Do you want to sort the summary?",
            ["no", "yes"],
            key="summary_sort_choice"
        )
        if sort_summary == "yes":
            summary_sort_column = st.selectbox(
                "Choose column to sort by",
                ["Total", "Average", "Highest"],
                key="summary_sort_column"
            )
            summary_sort_direction = st.selectbox(
                "Order",
                ["ascending", "descending"],
                key="summary_sort_direction"
            )
            summary_df = summary_df.sort_values(
                summary_sort_column,
                ascending=(summary_sort_direction == "ascending")
            )

        st.subheader("Summary Report")
        st.dataframe(summary_df)

    # ── Write to Excel in Memory ───────────────────────────────
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer) as writer:
        combined_df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        if summary_df is not None:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

    # ── Apply Bold Headers with openpyxl ───────────────────────
    excel_buffer.seek(0)
    workbook = load_workbook(excel_buffer)

    for cell in workbook["Cleaned Data"][1]:
        cell.font = Font(bold=True)

    if summary_df is not None:
        for cell in workbook["Summary"][1]:
            cell.font = Font(bold=True)

    # ── Save Final Workbook to Memory ──────────────────────────
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)

    # ── Show Stats and Download Button ─────────────────────────
    st.metric("Duplicates Removed", duplicates_removed)
    st.metric("Empty Cells Filled", empty_cell_count)
    st.download_button("Download Cleaned Excel", final_output.getvalue(), "cleaned.xlsx")

else:
    st.write("Please upload a file to get started.")
    

