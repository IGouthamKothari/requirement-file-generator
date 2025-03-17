import io
import os
import pandas as pd
import streamlit as st
import tempfile
import shutil

def extract_top_info(df):
    """Extracts job details from the first few rows of the Excel file."""
    job_number = df.iloc[2, 2] if pd.notna(df.iloc[2, 2]) else "Unknown"
    customer_name = df.iloc[1, 2] if pd.notna(df.iloc[1, 2]) else "Unknown"
    order_qty = df.iloc[4, 2] if pd.notna(df.iloc[4, 2]) else "Unknown"
    return job_number, customer_name, order_qty

def process_excel_file(file_path):
    """Reads and processes the new Excel format."""
    df = pd.read_excel(file_path, sheet_name=0, header=None)
    
    job_number, customer_name, order_qty = extract_top_info(df)
    
    # Find header row dynamically (we assume 'Sr. No.' appears in a specific column range)
    header_row_idx = None
    for i in range(len(df)):
        if "SR." in str(df.iloc[i, 0]).upper() and "MATERIAL" in str(df.iloc[i, 1]).upper():
            header_row_idx = i
            break
    
    if header_row_idx is None:
        return None, f"Header row not found in {file_path}"
    
    df_data = pd.read_excel(file_path, sheet_name=0, header=header_row_idx)
    df_data = df_data.dropna(how='all')  # Remove empty rows
    df_data = df_data.rename(columns=lambda x: str(x).strip())  # Clean column names
    
    # Ensure required columns exist
    required_columns = ["Material", "Specification", "Unit", "QTY/ UNIT", "NO. OF UNIT", "TOTAL QTY"]
    missing_columns = [col for col in required_columns if col not in df_data.columns]
    if missing_columns:
        return None, f"Missing columns {missing_columns} in {file_path}"
    
    # Add job info to data
    df_data["Job Number"] = job_number
    df_data["Customer Name"] = customer_name
    df_data["Order Qty"] = order_qty
    
    return df_data, None

def process_uploaded_files(uploaded_files):
    """Processes multiple uploaded Excel files and generates a combined report."""
    all_data = []
    errors = []
    
    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp.close()
            df, error = process_excel_file(tmp.name)
            if df is not None:
                all_data.append(df)
            if error:
                errors.append(error)
    
    if not all_data:
        return None, errors
    
    combined_df = pd.concat(all_data, ignore_index=True)
    output_excel = io.BytesIO()
    
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, sheet_name="Combined Data", index=False)
    
    output_excel.seek(0)
    return output_excel, errors

# Streamlit UI
st.title("New Excel Format Processor")
st.write("Upload Excel files to process them.")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("Process Files"):
    if not uploaded_files:
        st.error("Please upload at least one file.")
    else:
        output_excel_io, errors = process_uploaded_files(uploaded_files)
        
        if errors:
            st.subheader("Errors")
            st.text_area("Error Log", "\n".join(errors), height=200)
        
        if output_excel_io:
            st.success("Processing complete! Download the combined file below.")
            st.download_button(
                label="Download Combined Excel",
                data=output_excel_io,
                file_name="combined_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
