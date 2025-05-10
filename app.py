import streamlit as st
import pandas as pd
import numpy as np
import os

# Excel file path (update if needed)
EXCEL_FILE_PATH = r"C:\Users\GAURANG MHATRE\Downloads\IRIS_Public_Assignment-main\IRIS_Public_Assignment-main\Data\capbudg.xls"

@st.cache_data
def read_excel_file():
    """Read the Excel file and return a dictionary of DataFrames for each sheet."""
    try:
        excel_file = pd.ExcelFile(EXCEL_FILE_PATH)
        dataframes = {
            sheet: pd.read_excel(excel_file, sheet_name=sheet).replace({np.nan: None})
            for sheet in excel_file.sheet_names
        }
        return dataframes
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return {}

def main():
    st.set_page_config(page_title="Excel Table Viewer", layout="wide")
    st.title("📊 Excel Table Viewer and Row Sum Calculator")

    # Load data
    dataframes = read_excel_file()
    sheet_names = list(dataframes.keys())

    if not sheet_names:
        st.warning("No tables found in the Excel file.")
        return

    # Sheet selection
    selected_sheet = st.selectbox("Select a table (sheet):", sheet_names)
    df = dataframes[selected_sheet]

    if df.empty:
        st.warning(f"No data found in sheet: {selected_sheet}")
        return

    # Display the table
    st.subheader(f"Table Preview: {selected_sheet}")
    st.dataframe(df)

    # Row name selection (based on first column)
    row_names = [str(val) for val in df.iloc[:, 0].tolist() if val is not None]
    selected_row = st.selectbox("Select a row to calculate sum:", row_names)

    if selected_row:
        # Find the row index
        row_index = None
        for idx, val in enumerate(df.iloc[:, 0].tolist()):
            if str(val) == selected_row:
                row_index = idx
                break

        if row_index is not None:
            row_data = df.iloc[row_index, 1:].replace({None: 0})
            numeric_data = pd.to_numeric(row_data, errors='coerce').fillna(0)
            row_sum = numeric_data.sum()

            st.success(f"Sum of row '{selected_row}' is: {float(row_sum):,.2f}")
        else:
            st.error(f"Row '{selected_row}' not found in table '{selected_sheet}'")

if __name__ == "__main__":
    main()
