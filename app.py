import streamlit as st
import pandas as pd
import io

# Process Column O and P (skip header row)
def process_columns_OP(df):
    if df.shape[1] >= 16:
        for i in range(1, len(df)):
            if pd.notna(df.iat[i, 14]) and str(df.iat[i, 14]).strip() != "":
                df.iat[i, 14] = "United States"
            if pd.notna(df.iat[i, 15]) and str(df.iat[i, 15]).strip() != "":
                df.iat[i, 15] = "Home"
    return df

# Process each sheet in uploaded Excel file
def process_excel_file(file):
    xl = pd.ExcelFile(file)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)  # Don't treat first row as header
        df = process_columns_OP(df)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()
    output.seek(0)
    return output

# Streamlit interface
st.title("📄 Excel Formatter: Columns O and P")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    processed_file = process_excel_file(uploaded_file)
    st.success("Columns O and P processed successfully across all sheets.")
    st.download_button(
        label="📥 Download Processed File",
        data=processed_file,
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
