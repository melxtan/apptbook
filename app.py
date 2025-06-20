import streamlit as st
from excel_utils import process_files
import pandas as pd
from io import BytesIO

st.title("Daily Excel Macro Replacement")

csv_file = st.file_uploader("Upload CSV file", type="csv")
excel_file = st.file_uploader("Upload Excel file", type="xlsx")

if csv_file and excel_file:
    output = process_files(csv_file, excel_file)
    st.success("Processing complete. Download updated file below.")
    st.download_button("Download Updated Excel", output.getvalue(), file_name="Processed_Apptbook.xlsx")

    st.markdown("""
    <style>
    .css-1d391kg { color: black; }
    </style>
    """, unsafe_allow_html=True)
