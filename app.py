import io
import os
import re
import zipfile
import shutil
import tempfile
import pandas as pd
import streamlit as st

def create_sample_dataframe():
    """Creates a sample dataframe mimicking the structure of the provided Excel."""
    data = {
        "Sr.No.": [1, 2, 3, 4, 5],
        "Material": ["Copper Wire", "Aluminum Sheet", "Steel Rod", "Plastic Cover", "Rubber Seal"],
        "Specification": ["10mm thick", "5mm sheet", "2m long", "PVC Type A", "High-density"],
        "Units": ["KG", "KG", "Meter", "Piece", "Piece"],
        "Qty. per unit": [5, 10, 2, 1, 3],
        "No of units": [20, 15, 50, 100, 200],
        "Total Qty.": [100, 150, 100, 100, 600]
    }
    return pd.DataFrame(data)

def process_dataframe(df):
    """Applies logical operations similar to the original code."""
    df["Total Qty."] = df["Qty. per unit"] * df["No of units"]
    df["Material"] = df["Material"].astype(str).str.strip()
    df = df[df["Material"] != ""]
    df = df.groupby(["Material", "Units"], as_index=False)["Total Qty."].sum()
    return df

def process_excel_file():
    """Processes the generated sample data frame and prepares an Excel output."""
    df = create_sample_dataframe()
    processed_df = process_dataframe(df)
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        processed_df.to_excel(writer, sheet_name="ProcessedData", index=False)
    output_excel.seek(0)
    return output_excel, processed_df

st.title("Excel Processor - Custom Format")
if st.button("Generate Sample & Process"):
    output_excel_io, df_preview = process_excel_file()
    st.subheader("Processed Data Preview")
    st.dataframe(df_preview)
    
    st.download_button(
        label="Download Processed Excel",
        data=output_excel_io,
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
