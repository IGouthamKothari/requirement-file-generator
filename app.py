import io
import os
import re
import zipfile
import shutil
import tempfile
import base64
from typing import List, Tuple
import pandas as pd
import streamlit as st

# -------------------------
# Processing Functions
# -------------------------

def extract_top_info(excel_file, sheet_name, expected_headers, nrows=20):
    df_temp = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=nrows)
    job_number = None
    top_qty = 0
    client_name = None
    header_row_index = None
    for i in range(len(df_temp)):
        row_values = df_temp.iloc[i].astype(str).str.strip().tolist()
        row_str = " ".join(row_values).lower()
        # Detect job number.
        for cell in row_values:
            if isinstance(cell, str) and "job.no." in cell.lower():
                parts = cell.split(":")
                if len(parts) > 1:
                    job_number = parts[1].strip()
        # Detect top-level "Qty:" line.
        qty_match = re.search(r"qty:\s*(\d+)", row_str)
        if qty_match:
            top_qty = int(qty_match.group(1))
        # Detect client name.
        for cell in row_values:
            if isinstance(cell, str) and "customer:" in cell.lower():
                parts = cell.split(":")
                if len(parts) > 1:
                    client_name = parts[1].strip()
        # Detect header row.
        if all(h in row_values for h in expected_headers):
            header_row_index = i
    return job_number, top_qty, client_name, header_row_index

def clean_numeric(x):
    try:
        if isinstance(x, str):
            x = x.replace(",", "").strip()
        return float(x)
    except Exception:
        return 0.0

def merge_group(group):
    parent = group.iloc[0].copy()
    parent["Material"] = str(parent["Material"]).strip() if pd.notna(parent["Material"]) else ""
    for col in ["Qty. per unit", "No of units", "Total Qty."]:
        total = group[col].apply(clean_numeric).sum()
        parent[col] = total
    return parent

def merge_child_rows(df):
    merged_df = df.copy()
    merged_df["OriginalSrNo"] = pd.to_numeric(merged_df["Sr.No."], errors="coerce")
    merged_df["Sr.No."] = merged_df["Sr.No."].ffill()
    grouped = merged_df.groupby("Sr.No.", as_index=False)
    merged_rows = [merge_group(group.sort_index()) for name, group in grouped]
    result = pd.DataFrame(merged_rows)
    result.drop(columns=["OriginalSrNo"], inplace=True, errors="ignore")
    return result

def assign_group_key(spec):
    s = spec.strip()
    match = re.match(r"^([A-Za-z0-9]{2,})", s)
    return match.group(1).lower() if match else "other"

def validate_row(row, sheet_name, file_name, logs):
    errors = []
    if row["Total Qty."] == 0:
        errors.append("Total Qty. is zero")
    if not row["Units"] or str(row["Units"]).strip() == "":
        errors.append("Units is missing")
    if row["Qty. per unit"] == 0 or row["No of units"] == 0:
        errors.append("Numeric value is zero")
    if errors:
        logs.append(f"Validation error in file '{file_name}', sheet '{sheet_name}', row {row.name}: {', '.join(errors)}.")

def process_excel_files(input_files: List[Tuple[str, str]]) -> Tuple[io.BytesIO, str, List[str]]:
    """
    input_files: list of tuples (temp_file_path, original_file_name)
    Returns (output_excel as BytesIO, temporary directory, logs)
    """
    logs = []
    data_dir = tempfile.mkdtemp()
    output_file = os.path.join(data_dir, "combined_requirements.xlsx")
    expected_headers = [
        "Sr.No.",
        "Material",
        "Specifications",  # Will be renamed to "Units"
        "Qty. per unit",
        "No of units",
        "Total Qty.",
        "Remarks"
    ]
    all_data = []
    all_top_qty = 0
    job_to_client = {}

    for file_path, original_file in input_files:
        for sheet_name in pd.ExcelFile(file_path).sheet_names:
            job_number, top_qty, client_name, hdr_idx = extract_top_info(
                excel_file=file_path,
                sheet_name=sheet_name,
                expected_headers=expected_headers,
                nrows=20
            )
            if hdr_idx is None:
                continue
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=hdr_idx)
            except Exception:
                continue
            df.columns = df.columns.str.strip()
            if not set(expected_headers).issubset(df.columns):
                continue
            needed_df = df[expected_headers].copy()
            needed_df.rename(columns={"Specifications": "Units"}, inplace=True)
            merged_df = merge_child_rows(needed_df)
            for col in ["Qty. per unit", "No of units", "Total Qty."]:
                merged_df[col] = merged_df[col].apply(clean_numeric)
            for idx, row in merged_df.iterrows():
                validate_row(row, sheet_name, original_file, logs)
            merged_df["Material"] = merged_df["Material"].astype(str).str.strip()
            merged_df = merged_df[merged_df["Material"] != ""]
            mask_all_zero = (
                (merged_df["Qty. per unit"] == 0) &
                (merged_df["No of units"] == 0) &
                (merged_df["Total Qty."] == 0)
            )
            merged_df = merged_df[~mask_all_zero]
            if not job_number:
                job_number = original_file.replace("WORK ORDER ", "").replace(".xlsx", "").strip()
            all_top_qty += top_qty
            job_to_client[job_number] = client_name or "UNKNOWN_CLIENT"
            merged_df["UniqueOrderID"] = job_number + "_" + sheet_name
            all_data.append(merged_df)
    if not all_data:
        logs.append("No data found in any sheet.")
        combined_df = pd.DataFrame()
    else:
        combined_df = pd.concat(all_data, ignore_index=True)
    try:
        overall_df = (
            combined_df
            .groupby(["Material", "Units"], as_index=False)["Total Qty."]
            .sum()
            .rename(columns={"Material": "Specification", "Total Qty.": "Total_Qty"})
        )
        overall_df["GroupKey"] = overall_df["Specification"].apply(assign_group_key)
        overall_df = overall_df.sort_values(["GroupKey", "Specification"]).reset_index(drop=True)
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            workbook = writer.book
            bold_format = workbook.add_format({"bold": True})
            ws_overall = workbook.add_worksheet("OverallRequirement")
            writer.sheets["OverallRequirement"] = ws_overall
            header = ["Sl. No.", "Specification", "Units", "Total_Qty"]
            ws_overall.write_row(0, 0, header)
            current_row = 1
            sl_no = 1
            current_group = None
            for idx, row in overall_df.iterrows():
                group_key = row["GroupKey"]
                if current_group != group_key:
                    ws_overall.write(current_row, 0, group_key.upper(), bold_format)
                    current_row += 1
                    current_group = group_key
                ws_overall.write(current_row, 0, sl_no)
                ws_overall.write(current_row, 1, row["Specification"])
                ws_overall.write(current_row, 2, row["Units"])
                ws_overall.write(current_row, 3, row["Total_Qty"])
                current_row += 1
                sl_no += 1
            current_row += 1
            ws_overall.write(current_row, 0, "TOP-LEVEL QTY SUM:", bold_format)
            ws_overall.write(current_row, 1, all_top_qty)
            current_row += 2
            ws_overall.write(current_row, 0, "JOB NUMBER", bold_format)
            ws_overall.write(current_row, 1, "CLIENT NAME", bold_format)
            current_row += 1
            for jn in sorted(job_to_client.keys()):
                ws_overall.write(current_row, 0, jn)
                ws_overall.write(current_row, 1, job_to_client[jn])
                current_row += 1

            per_order_df = (
                combined_df
                .groupby(["UniqueOrderID", "Material", "Units"], as_index=False)["Total Qty."]
                .sum()
                .rename(columns={"Material": "Specification", "Total Qty.": "Total_Qty"})
            )
            per_order_df.insert(0, "Sl. No.", range(1, len(per_order_df) + 1))
            per_order_df = per_order_df[["Sl. No.", "UniqueOrderID", "Specification", "Units", "Total_Qty"]]
            per_order_df.to_excel(writer, sheet_name="PerOrderRequirement", index=False)
        output_excel.seek(0)
    except Exception:
        logs.append("Error creating output Excel file.")
        output_excel = io.BytesIO()
    return output_excel, data_dir, logs

# -------------------------
# Streamlit Frontend
# -------------------------

st.title("Excel Processor")
st.write("Upload multiple Excel files to process them. The output will include a combined Excel file, a preview, and a list of validation errors.")

# Use session state to disable the process button once clicked
if "processing" not in st.session_state:
    st.session_state.processing = False

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("Process Files", disabled=st.session_state.processing):
    st.session_state.processing = True
    with st.spinner("Processing files..."):
        if not uploaded_files:
            st.error("Please upload at least one file.")
        else:
            file_info = []  # List of tuples: (temp_file_path, original_file_name)
            for file in uploaded_files:
                try:
                    file_bytes = file.read()
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    tmp.write(file_bytes)
                    tmp.close()
                    file_info.append((tmp.name, file.name))
                except Exception as e:
                    st.error(f"Error saving file {file.name}: {e}")
            output_excel_io, work_dir, logs = process_excel_files(file_info)
            
            # Show validation errors in a scrollable text area
            st.subheader("Validation Errors")
            if logs:
                st.text_area("Errors", "\n".join(logs), height=200)
            else:
                st.success("No validation errors.")
            
            # Display a preview of the combined Excel output (OverallRequirement sheet)
            try:
                output_excel_io.seek(0)
                df_preview = pd.read_excel(output_excel_io, sheet_name="OverallRequirement")
                st.subheader("Combined Excel Output (OverallRequirement)")
                st.dataframe(df_preview)
            except Exception as e:
                st.error(f"Error reading Excel output: {e}")
            
            # Reset pointer for download
            output_excel_io.seek(0)
            
            # Create a ZIP file containing both the Excel output and the validation errors
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("combined_requirements.xlsx", output_excel_io.getvalue())
                error_text = "\n".join(logs) if logs else "No validation errors."
                zip_file.writestr("validation_errors.txt", error_text)
            zip_buffer.seek(0)
            
            # Download buttons
            st.download_button(
                label="Download Combined Excel",
                data=output_excel_io,
                file_name="combined_requirements.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Processed ZIP",
                data=zip_buffer,
                file_name="processed_output.zip",
                mime="application/zip"
            )
            
            # Clean up temporary files.
            for tmp_path, _ in file_info:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            try:
                shutil.rmtree(work_dir, ignore_errors=True)
            except Exception:
                pass
    st.session_state.processing = False
