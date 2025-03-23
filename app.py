import io
import os
import re
import zipfile
import shutil
import tempfile
import pandas as pd
import streamlit as st

def extract_top_info(df):
    """Extracts job number, client name, and order quantity from the first few rows."""
    job_number = None
    client_name = None
    top_qty = 0
    for i in range(len(df)):
        row_values = df.iloc[i].astype(str).str.strip().tolist()
        row_str = " ".join(row_values).lower()
        if "job no." in row_str:
            job_number = row_values[-1]
        if "customer" in row_str:
            client_name = row_values[-1]
        if "order qty." in row_str:
            qty_match = re.search(r"(\d+)", row_str)
            if qty_match:
                top_qty = int(qty_match.group(1))
    return job_number, client_name, top_qty

def process_excel(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    
    job_number, client_name, top_qty = extract_top_info(df)
    
    header_row_index = 7  # Adjusted for the new format
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_index)
    
    expected_columns = {
        "SR. NO.": "Sr.No.",
        "MATERIAL": "Material",
        "SPECIFICATION": "Specification",
        "UNIT": "Units",
        "QTY/ UNIT": "Qty. per unit",
        "TOTAL QTY": "Total Qty."
    }
    
    df.rename(columns=expected_columns, inplace=True)
    df = df[list(expected_columns.values())]  # Keep only relevant columns
    
    df.dropna(subset=["Material"], inplace=True)
    df["Total Qty."] = df["Total Qty."].apply(pd.to_numeric, errors="coerce").fillna(0)
    
    return df, job_number, client_name, top_qty

st.title("Supertherm Excel Processor")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    
    df_processed, job_no, client, top_qty = process_excel(tmp_path)
    
    st.subheader("Processed Data Preview")
    st.dataframe(df_processed)
    
    st.write(f"**Job Number:** {job_no}")
    st.write(f"**Client Name:** {client}")
    st.write(f"**Total Order Quantity:** {top_qty}")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_processed.to_excel(writer, sheet_name="ProcessedData", index=False)
    output.seek(0)
    
    st.download_button(
        label="Download Processed Excel",
        data=output,
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    os.remove(tmp_path)
