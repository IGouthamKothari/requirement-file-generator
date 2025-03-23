import io
import os
import pandas as pd
import streamlit as st

def compare_with_tally(combined_df, tally_df):
    """
    Compares the combined requirement sheet with the tally stock data.
    Returns a DataFrame with the comparison results.
    """
    comparison_results = []
    
    for _, row in combined_df.iterrows():
        material = row["Specification"].strip()
        required_qty = row["Total_Qty"]
        units = row["Units"]
        
        # Find material in tally stock
        stock_row = tally_df[tally_df["Material Name"].str.strip().str.lower() == material.lower()]
        
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
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
tally_file = st.file_uploader("Upload Tally Stock Excel", type=["xlsx", "xls"], accept_multiple_files=False)

if st.button("Process and Compare"):
    if not uploaded_files or not tally_file:
        st.error("Please upload both requirement and tally stock files.")
    else:
        st.spinner("Processing files...")
        file_info = []
        for file in uploaded_files:
            file_bytes = file.read()
            temp_path = f"/tmp/{file.name}"  # Temporary path
            with open(temp_path, "wb") as f:
                f.write(file_bytes)
            file_info.append(temp_path)
        
        # Process the uploaded Excel files to generate combined requirements (mock function here)
        combined_df = pd.DataFrame({  # Replace with actual processing function
            "Specification": ["Material A", "Material B"],
            "Units": ["kg", "pcs"],
            "Total_Qty": [100, 200]
        })
        
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
