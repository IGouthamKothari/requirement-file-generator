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
    combined_df["Total Qty."] = combined_df["Qty. per unit"] * combined_df["No of units"]
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
    output_excel_io, df_preview = process_uploaded_files(uploaded_files)
    if output_excel_io:
        st.subheader("Processed Data Preview")
        st.dataframe(df_preview)
        
        st.download_button(
            label="Download Processed Excel",
            data=output_excel_io,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No valid data found in the uploaded files.")
