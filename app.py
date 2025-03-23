import io
import pandas as pd
import streamlit as st

def load_and_merge_requirements(files):
    """Loads multiple Excel files and merges them into a single DataFrame."""
    all_dfs = []
    for file in files:
        df = pd.read_excel(file)
        all_dfs.append(df)
    return pd.concat(all_dfs, ignore_index=True)

# Streamlit Frontend
st.title("Excel Processor - Combined Sheet Generator")
uploaded_files = st.file_uploader("Upload Requirement Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("Generate Combined Sheet"):
    if not uploaded_files:
        st.error("Please upload requirement files.")
    else:
        with st.spinner("Processing files..."):
            # Load and merge requirement Excel files
            combined_df = load_and_merge_requirements(uploaded_files)
            
            # Save to Excel
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                combined_df.to_excel(writer, sheet_name="OverallRequirement", index=False)
            output_excel.seek(0)
            
            # Provide download button
            st.download_button(
                label="Download Combined Excel",
                data=output_excel,
                file_name="combined_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing complete! Download the output file above.")
