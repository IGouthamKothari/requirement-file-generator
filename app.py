import io
import os
import pandas as pd
import streamlit as st

def load_and_merge_requirements(files):
    """Loads multiple Excel files and merges them into a single DataFrame."""
    all_dfs = []
    for file in files:
        df = pd.read_excel(file)
        all_dfs.append(df)
    return pd.concat(all_dfs, ignore_index=True)

def compare_with_tally(combined_df, tally_df):
    """
    Compares the combined requirement sheet with the tally stock data.
    Returns a DataFrame with the comparison results.
    """
    comparison_results = []
    
    for _, row in combined_df.iterrows():
        material = str(row["Specification"]).strip()
        required_qty = row["Total_Qty"]
        units = row["Units"]
        
        # Ensure material name comparison is case-insensitive and whitespace-trimmed
        stock_row = tally_df[tally_df["Material Name"].astype(str).str.strip().str.lower() == material.lower()]
        
        if not stock_row.empty:
            available_stock = stock_row.iloc[0]["Stock Quantity"]
            stock_difference = available_stock - required_qty
            status = "Exact Match" if stock_difference == 0 else ("Surplus" if stock_difference > 0 else "Shortage")
        else:
            available_stock = "Not Found"
            stock_difference = "N/A"
            status = "Not Found"
        
        comparison_results.append([material, units, required_qty, available_stock, stock_difference, status])
    
    comparison_df = pd.DataFrame(comparison_results, columns=["Specification", "Units", "Required Qty", "Available Stock", "Stock Difference", "Status"])
    return comparison_df

# Streamlit Frontend
st.title("Excel Processor with Stock Comparison")
uploaded_files = st.file_uploader("Upload Requirement Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
tally_file = st.file_uploader("Upload Tally Stock Excel", type=["xlsx", "xls"], accept_multiple_files=False)

if st.button("Process and Compare"):
    if not uploaded_files or not tally_file:
        st.error("Please upload both requirement and tally stock files.")
    else:
        with st.spinner("Processing files..."):
            # Load and merge requirement Excel files
            combined_df = load_and_merge_requirements(uploaded_files)
            
            # Load tally stock file
            tally_df = pd.read_excel(tally_file)
            
            # Compare with tally stock
            comparison_df = compare_with_tally(combined_df, tally_df)
            
            # Save to Excel
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                combined_df.to_excel(writer, sheet_name="OverallRequirement", index=False)
                comparison_df.to_excel(writer, sheet_name="Stock Comparison", index=False)
            output_excel.seek(0)
            
            # Provide download button
            st.download_button(
                label="Download Processed Excel",
                data=output_excel,
                file_name="processed_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing complete! Download the output file above.")
