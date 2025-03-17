import io
import os
import re
import zipfile
import shutil
import tempfile
import pandas as pd
import streamlit as st

def extract_data_from_excel(excel_file):
    """Extracts data from the new Excel format."""
    df_temp = pd.read_excel(excel_file, sheet_name=0, header=[7, 8])  # Read multi-row header
    
    # Rename columns to expected format
    column_mapping = {
        ('SR.', 'NO.'): 'Sr.No.',
        ('SPECIFICATION', '\xa0'): 'Material',
        ('UNIT', '\xa0'): 'Units',
        ('QTY/', 'UNIT'): 'Qty. per unit',
        ('NO. OF', 'UNIT'): 'No of units',
        ('TOTAL', 'QTY'): 'Total Qty.'
    }
    
    df_temp.rename(columns=column_mapping, inplace=True)
    df_temp = df_temp[list(column_mapping.values())]  # Keep only required columns
    
    # Convert numeric columns
    for col in ['Qty. per unit', 'No of units', 'Total Qty.']:
        df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0)
    
    # Remove empty material rows
    df_temp = df_temp[df_temp['Material'].notna()]
    df_temp['Material'] = df_temp['Material'].astype(str).str.strip()
    
    return df_temp

def process_excel_files(input_files):
    logs = []
    output_excel = io.BytesIO()
    all_data = []
    
    for file_path, original_file in input_files:
        try:
            df = extract_data_from_excel(file_path)
            all_data.append(df)
        except Exception as e:
            logs.append(f"Error processing {original_file}: {str(e)}")
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df.to_excel(output_excel, index=False, sheet_name='ProcessedData')
    
    output_excel.seek(0)
    return output_excel, logs

# Streamlit Frontend
st.title("Excel Processor - New Format")
st.write("Upload Excel files to process them.")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("Process Files"):
    if not uploaded_files:
        st.error("Please upload at least one file.")
    else:
        file_info = []
        for file in uploaded_files:
            file_bytes = file.read()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp.write(file_bytes)
            tmp.close()
            file_info.append((tmp.name, file.name))
        
        output_excel_io, logs = process_excel_files(file_info)
        
        st.subheader("Validation Errors")
        if logs:
            st.text_area("Errors", "\n".join(logs), height=200)
        else:
            st.success("No validation errors.")
        
        output_excel_io.seek(0)
        df_preview = pd.read_excel(output_excel_io, sheet_name='ProcessedData')
        st.subheader("Processed Data Preview")
        st.dataframe(df_preview)
        
        st.download_button(
            label="Download Processed Excel",
            data=output_excel_io,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        for tmp_path, _ in file_info:
            os.remove(tmp_path)
