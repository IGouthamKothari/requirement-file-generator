import io
import os
import pandas as pd
import streamlit as st
import tempfile
import shutil

def process_uploaded_files(uploaded_files):
    """Processes multiple uploaded Excel files and generates a combined report."""
    all_data = []
    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name
        
        xls = pd.ExcelFile(tmp_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            all_data.append(df)
        os.remove(tmp_path)
    
    if not all_data:
        return None
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Standardize column names
    combined_df.columns = combined_df.columns.str.strip().str.lower()
    
    # Debugging: Print available columns
    print("Columns in DataFrame:", combined_df.columns.tolist())
    
    # Dynamically detect relevant columns
    qty_col = [col for col in combined_df.columns if "qty" in col and "per unit" in col]
    no_units_col = [col for col in combined_df.columns if "no" in col and "unit" in col]
    
    if qty_col and no_units_col:
        qty_col = qty_col[0]  # Use the first matching column
        no_units_col = no_units_col[0]
        combined_df["Total Qty."] = combined_df[qty_col] * combined_df[no_units_col]
    else:
        raise ValueError("Required columns for calculation not found in the uploaded files.")
    
    combined_df["Material"] = combined_df["Material"].astype(str).str.strip()
    combined_df = combined_df[combined_df["Material"] != ""]
    combined_df = combined_df.groupby(["Material", "Units"], as_index=False)["Total Qty."].sum()
    
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, sheet_name="ProcessedData", index=False)
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
