import streamlit as st
import pandas as pd
import io

# Process a DataFrame to modify column O (15th column) from row 2 downward
def process_column_O(df):
    if df.shape[1] >= 15:  # Ensure column O exists (index 14)
        for i in range(1, len(df)):  # Skip header row (row 0)
            cell = df.iat[i, 14]  # Column O = index 14
            if pd.notna(cell) and str(cell).strip() != "":
                df.iat[i, 14] = "United States"
    return df

# Process every sheet in the Excel file
def process_excel_file(file):
    xl = pd.ExcelFile(file)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)  # Read without assuming header
        df = process_column_O(df)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()
    output.seek(0)
    return output

# Streamlit UI
st.title("📄 Excel Column O Formatter (All Sheets)")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    processed_file = process_excel_file(uploaded_file)
    st.success("Column O processed successfully across all sheets.")
    st.download_button(
        label="📥 Download Processed File",
        data=processed_file,
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
