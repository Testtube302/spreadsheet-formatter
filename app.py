import streamlit as st
import pandas as pd
import io

# Function to apply formatting rules to a DataFrame
def process_columns(df):
    for i in range(1, len(df)):  # Skip header row (row 0)
        # Column O = 14
        if df.shape[1] > 14 and pd.notna(df.iat[i, 14]) and str(df.iat[i, 14]).strip() != "":
            df.iat[i, 14] = "United States"

        # Column P = 15
        if df.shape[1] > 15 and pd.notna(df.iat[i, 15]) and str(df.iat[i, 15]).strip() != "":
            df.iat[i, 15] = "Home"

        # Columns S, T, U = 18, 19, 20
        for col_index in [18, 19, 20]:
            if df.shape[1] > col_index:
                val = str(df.iat[i, col_index]).strip().lower() if pd.notna(df.iat[i, col_index]) else ""
                if val in ["yes", "y"]:
                    df.iat[i, col_index] = "true"
                elif val in ["no", "n"]:
                    df.iat[i, col_index] = "false"

        # Column Y = 24
        if df.shape[1] > 24 and pd.notna(df.iat[i, 24]) and str(df.iat[i, 24]).strip() != "":
            df.iat[i, 24] = "Home"

        # Columns AB, AD, AE = 27, 29, 30 → Clear if data exists
        for col_index in [27, 29, 30]:
            if df.shape[1] > col_index:
                cell_value = df.iat[i, col_index]
                if pd.notna(cell_value) and str(cell_value).strip() != "":
                    df.iat[i, col_index] = ""

    return df

# Process all sheets in the Excel file
def process_excel_file(file):
    xl = pd.ExcelFile(file)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)  # No header assumed
        df = process_columns(df)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()
    output.seek(0)
    return output

# Streamlit UI
st.title("📊 Excel Multi-Sheet Formatter")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    processed_file = process_excel_file(uploaded_file)
    st.success("✅ File processed: Columns O, P, S, T, U, Y, AB, AD, AE updated.")
    st.download_button(
        label="📥 Download Processed File",
        data=processed_file,
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
