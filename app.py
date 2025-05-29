import streamlit as st
import pandas as pd
import io

def process_dataframe(df):
    df.columns = df.columns.str.strip()

    # Remove rows A to E (first 5 rows)
    df = df.iloc[5:].reset_index(drop=True)

    col_map = {
        'O': 'United States',
        'P': 'Home',
        'Y': 'Home',
        'AH': 'United States',
        'AI': 'Florida',
    }

    for col, val in col_map.items():
        if col in df.columns:
            df[col] = df[col].where(df[col].isna(), val)

    # Boolean transformation
    bool_map = {'yes': True, 'y': True, 'no': False, 'n': False}
    for col in ['S', 'T', 'U'] + [chr(i) for i in range(ord('AK'), ord('AQ')+1)]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower().map(bool_map).fillna(df[col])

    # Clear content for specific columns
    for col in ['AB', 'AD', 'AE', 'AQ', 'AR']:
        if col in df.columns:
            df[col] = ""

    return df

def process_excel(file):
    xl = pd.ExcelFile(file)
    output = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df = process_dataframe(df)
        df.to_excel(output, sheet_name=sheet, index=False)

    output.save()
    return 'output.xlsx'

st.title("Excel/CSV Data Formatter")

uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        df = process_dataframe(df)

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button("Download Processed File", output, file_name="formatted_output.xlsx")
    else:
        output_path = process_excel(uploaded_file)
        with open(output_path, "rb") as f:
            st.download_button("Download Processed File", f, file_name="formatted_output.xlsx")
