import streamlit as st
import pandas as pd
import io
from string import ascii_uppercase

# Helper to generate Excel-style column labels (e.g., AK to AP)
def get_column_range(start, end):
    cols = []
    found = False
    for c1 in ascii_uppercase:
        for c2 in ascii_uppercase:
            col = c1 + c2
            if col == start:
                found = True
            if found:
                cols.append(col)
            if col == end:
                return cols
    return cols

def process_dataframe(df):
    df.columns = df.columns.str.strip()

    # Remove rows A to E (first 5 rows)
    df = df.iloc[5:].reset_index(drop=True)

    # Set specific values if cell is not empty
    replacements = {
        'O': 'United States',
        'P': 'Home',
        'Y': 'Home',
        'AH': 'United States',
        'AI': 'Florida'
    }

    for col, val in replacements.items():
        if col in df.columns:
            df[col] = df[col].where(df[col].isna(), val)

    # Boolean transformation: S, T, U and AK–AP
    bool_map = {'yes': True, 'y': True, 'no': False, 'n': False}
    boolean_columns = ['S', 'T', 'U'] + get_column_range('AK', 'AP')
    for col in boolean_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower().map(bool_map).fillna(df[col])

    # Clear specific columns
    for col in ['AB', 'AD', 'AE', 'AQ', 'AR']:
        if col in df.columns:
            df[col] = ""

    return df

def process_excel(file):
    xl = pd.ExcelFile(file)
    output_buffer = io.BytesIO()
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')

    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df = process_dataframe(df)
        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()
    output_buffer.seek(0)
    return output_buffer

# Streamlit UI
st.title("🧾 Spreadsheet Formatter")

uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        df = process_dataframe(df)

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button("📥 Download Formatted File", output, file_name="formatted_output.xlsx")

    else:
        output = process_excel(uploaded_file)
        st.download_button("📥 Download Formatted File", output, file_name="formatted_output.xlsx")
