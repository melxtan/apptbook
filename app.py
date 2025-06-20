import streamlit as st
from excel_utils import process_files
from redcap import fetch_redcap_data, parse_redcap_to_df, filter_new_records, update_mrn_sheet
import pandas as pd

st.title("Daily Excel Macro Replacement")

csv_file = st.file_uploader("Upload CSV file", type="csv")
excel_file = st.file_uploader("Upload Excel file", type="xlsx")

if csv_file and excel_file:
    output = process_files(csv_file, excel_file)
    st.success("Processing complete. Download updated file below.")
    st.download_button("Download Updated Excel", output.getvalue(), file_name="Processed_Apptbook.xlsx")

st.header("REDCap MRN Data Integration")

use_redcap = st.checkbox("Fetch and update MRN sheet from REDCap")
if use_redcap:
    api_key = st.secrets["REDCAP"]["KEY_1"]
    try:
        redcap_data = fetch_redcap_data(api_key)
        df_redcap = parse_redcap_to_df(redcap_data)

        st.write("Fetched REDCap Records:", df_redcap.head())

        uploaded_mrn = st.file_uploader("Upload Existing MRN Sheet", type="xlsx")
        if uploaded_mrn:
            xl = pd.ExcelFile(uploaded_mrn)
            existing_mrn_df = xl.parse("MRN", header=0)
            new_records = filter_new_records(df_redcap, existing_mrn_df)
            updated_mrn_df = update_mrn_sheet(existing_mrn_df, new_records)

            output_mrn = updated_mrn_df.to_excel(index=False, engine="openpyxl")
            st.download_button("Download Updated MRN Sheet", output_mrn, file_name="Updated_MRN.xlsx")
    except Exception as e:
        st.error(f"Error fetching REDCap data: {e}")
