import streamlit as st
import pandas as pd
import requests
import json
import os

# Load REDCap API keys from st.secrets
API_KEYS = [
    st.secrets["redcap"]["api_key_1"],
    st.secrets["redcap"]["api_key_2"]
]

REDCAP_URL = "https://redcap.med.usc.edu/api/"

st.set_page_config(page_title="REDCap Outpatient Processor", layout="wide")
st.title("Unified REDCap + Excel Processing App")

uploaded_file = st.file_uploader("Upload your Excel workbook:", type=[".xlsx"])
sharepoint_file = st.file_uploader("Upload SharePoint Excel file (for Notes update):", type="xlsx", key="sharepoint")

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    df_combined = None
    df_updated_notes = None
    df_processed_op = None

    # --- Refresh REDCap MRN Data ---
    if "MRN" in xl.sheet_names:
        df_existing = xl.parse("MRN")
        if "Full Case ID" in df_existing.columns:
            existing_case_ids = df_existing['Full Case ID'].astype(str).str.strip().dropna().unique().tolist()
            new_records = []

            for key in API_KEYS:
                data = {
                    'token': key,
                    'content': 'record',
                    'format': 'json',
                    'type': 'flat',
                    'rawOrLabel': 'label'
                }
                response = requests.post(REDCAP_URL, data=data)
                if response.status_code == 200:
                    records = response.json()
                    for record in records:
                        full_case_id = str(record.get("full_case_id", "")).strip()
                        mrn = record.get("mrn", "")
                        if mrn and full_case_id not in existing_case_ids:
                            new_records.append(record)
                else:
                    st.error(f"Failed REDCap pull: {response.status_code}")

            if new_records:
                df_new = pd.DataFrame(new_records)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True).drop_duplicates(subset=["mrn", "case_id", "full_case_id"])
                st.success(f"Imported {len(df_new)} new MRNs.")
            else:
                df_combined = df_existing.copy()
                st.info("No new MRNs found.")
        else:
            st.warning("'Full Case ID' column not found in MRN sheet.")

    # --- Update Notes Column W ---
    if sharepoint_file and "New OP" in xl.sheet_names:
        xl_sharepoint = pd.ExcelFile(sharepoint_file)
        if "New OP" in xl_sharepoint.sheet_names:
            df_local = xl.parse("New OP")
            df_share = xl_sharepoint.parse("New OP")
            df_local["W"] = df_share["W"]
            df_updated_notes = df_local.copy()
            st.success("Notes column W updated from SharePoint file.")

    # --- Process Outpatient Sheets ---
    required_sheets = ["New OP", "Routine", "MRN"]
    if all(sheet in xl.sheet_names for sheet in required_sheets):
        df_newop = xl.parse("New OP")
        df_routine = xl.parse("Routine").iloc[5:]  # from row 6
        df_mrn_lookup = xl.parse("MRN")

        df_routine = df_routine[df_routine.dropna(how='all').index]
        df_routine['color'] = 'red'
        df_newop['color'] = 'black'
        df_merged = pd.concat([df_newop, df_routine], ignore_index=True)

        for col in ['D', 'E', 'H']:
            if col in df_merged.columns:
                df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce')

        df_merged.sort_values(by=['D', 'E'], inplace=True)
        df_merged.drop_duplicates(subset=df_merged.columns[:26], inplace=True)

        df_merged['X'] = df_merged['B'].map(df_mrn_lookup.set_index('A')['B'])
        df_merged['Z'] = df_merged['B'].map(df_mrn_lookup.set_index('A')['F'])
        df_merged['AA'] = df_merged['B'].map(df_mrn_lookup.set_index('A')['G'])

        df_processed_op = df_merged.copy()
        st.success("Outpatient data processed and merged with Routine.")

    # --- Export everything ---
    if df_combined is not None or df_updated_notes is not None or df_processed_op is not None:
        with pd.ExcelWriter("Combined_Processed_File.xlsx", engine="xlsxwriter") as writer:
            if df_combined is not None:
                df_combined.to_excel(writer, index=False, sheet_name="MRN")
            if df_processed_op is not None:
                df_processed_op.to_excel(writer, index=False, sheet_name="New OP")
            elif df_updated_notes is not None:
                df_updated_notes.to_excel(writer, index=False, sheet_name="New OP")

        with open("Combined_Processed_File.xlsx", "rb") as f:
            st.download_button("Download Combined Processed File", f, file_name="Combined_Processed_File.xlsx")
else:
    st.info("Please upload your Excel workbook to begin.")
