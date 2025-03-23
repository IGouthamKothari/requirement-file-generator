import io
import os
import pandas as pd
import streamlit as st
import tempfile
import shutil
import re

def extract_top_info(df):
    """Extracts job number, customer name, and order quantity from the first few rows."""
    job_number = None
    customer_name = None
    order_qty = None
    for i in range(len(df)):
        row_values = df.iloc[i].astype(str).str.strip().tolist()
        row_str = " ".join(row_values).lower()
        if "job no." in row_str:
            job_number = row_values[-1]
        if "customer" in row_str:
            customer_name = row_values[-1]
        if "order qty." in row_str:
            qty_match = re.search(r"(\d+[+\d+]*)", row_str)
            if qty_match:
                order_qty = qty_match.group(1)
    return job_number, customer_name, order_qty

def process_uploaded_files(uploaded_files):
    """Processes multiple uploaded Excel files and generates a combined report."""
    all_data = []
    job_info = []
    
    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name
        
        xls = pd.ExcelFile(tmp_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            job_number, customer_name, order_qty = extract_top_info(df)
            df = pd.read_excel(xls, sheet_name=sheet_name, header=7)  # Adjust header row index
            job_info.append((job_number, customer_name, order_qty))
            all_data.append(df)
        os.remove(tmp_path)
    
    if not all_data:
        return None, None
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Standardize column names
    combined_df.columns = combined_df.columns.str.strip().str.lower()
    
    # Debugging: Print available columns
    print("Columns in DataFrame:", combined_df.columns.tolist())
    
    # Dynamically detect relevant columns
    qty_col = [col for col in combined_df.columns if "qty" in col and "per unit" in col]
    no_units_col = [col for col in combined_df.columns if "no" in col and "unit" in col]
    total_qty_col = [col for col in combined_df.columns if "total" in col and "qty" in col]
    
    if qty_col and no_units_col and total_qty_col:
        qty_col = qty_col[0]  # Use the first matching column
        no_units_col = no_units_col[0]
        total_qty_col = total_qty_col[0]
        combined_df["Total Qty."] = combined_df[total_qty_col]
    else:
        raise ValueError("Required columns for calculation not found in the uploaded files.")
    
    combined_df["Material"] = combined_df["Material"].astype(str).str.strip()
    combined_df = combined_df[combined_df["Material"] != ""]
    combined_df = combined_df.groupby(["Material", "units"], as_index=False)["Total Qty."].sum()
    
    # Create an Excel output
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, sheet_name="ProcessedData", index=False)
        pd.DataFrame(job_info, columns=["Job Number", "Customer", "Order Quantity"]).to_excel(writer, sheet_name="JobInfo", index=False)
    output_excel.seek(0)
    return output_excel, combined_df

st.title("Excel Processor - Multiple Files")
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files and st.button("Process Files"):
    try:
        output_excel_io, df_preview = process_uploaded_files(uploaded_files)
        st.subheader("Processed Data Preview")
        st.dataframe(df_preview)
        
        st.download_button(
            label="Download Processed Excel",
            data=output_excel_io,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error: {e}")
