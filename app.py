import streamlit as st
import pandas as pd
import io

# Process a single DataFrame to apply the column O rule
def process_column_O(df):
    if 'O' in df.columns:
        # Replace all non-empty values in column O, excluding the header
        df['O'] = df['O'].apply(lambda x: "United States" if pd.notna(x) and str(x).strip() != "" else "")
    return df

# Process all sheets in an uploaded Excel file
def process_excel_file(file):
    # Read the uploaded Excel file
    xl = pd.ExcelFile(file)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        df = process_column_O(df)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()
    output.seek(0)
    return output

# Streamlit UI
st.title("📄 Excel Column O Formatter")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    processed_file = process_excel_file(uploaded_file)
    st.success("File processed successfully.")
    st.download_button(
        label="📥 Download Processed File",
        data=processed_file,
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
