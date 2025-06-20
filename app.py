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

# Sidebar task selection
task = st.sidebar.selectbox("Choose Task", [
    "Refresh REDCap Data",
    "Process Outpatient Sheets",
    "Sync with SharePoint",
    "Update Notes Column W"
])

uploaded_file = st.file_uploader("Upload your Excel workbook:", type=[".xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)

    if task == "Refresh REDCap Data":
        if "MRN" not in xl.sheet_names:
            st.error("'MRN' sheet not found in the workbook.")
        else:
            df_existing = xl.parse("MRN")
            existing_case_ids = df_existing['full_case_id'].astype(str).str.strip().dropna().unique().tolist()
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
                st.download_button("Download Updated MRN Sheet", df_combined.to_excel(index=False), file_name="Updated_MRN.xlsx")
            else:
                st.info("No new MRNs found.")

    elif task == "Process Outpatient Sheets":
        required_sheets = ["New OP", "Routine", "MRN"]
        if not all(sheet in xl.sheet_names for sheet in required_sheets):
            st.error("Workbook must contain sheets: New OP, Routine, MRN")
        else:
            df_newop = xl.parse("New OP")
            df_routine = xl.parse("Routine").iloc[5:]  # From row 6
            df_mrn = xl.parse("MRN")

            df_routine = df_routine[df_routine.dropna(how='all').index]
            df_routine['color'] = 'red'
            df_newop['color'] = 'black'

            df_merged = pd.concat([df_newop, df_routine], ignore_index=True)
            for col in ['D', 'E', 'H']:
                if col in df_merged.columns:
                    df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce')
            df_merged.sort_values(by=['D', 'E'], inplace=True)
            df_merged.drop_duplicates(subset=df_merged.columns[:26], inplace=True)

            # Simulate VLOOKUP
            df_merged['X'] = df_merged['B'].map(df_mrn.set_index('A')['B'])
            df_merged['Z'] = df_merged['B'].map(df_mrn.set_index('A')['F'])
            df_merged['AA'] = df_merged['B'].map(df_mrn.set_index('A')['G'])

            with pd.ExcelWriter("Processed_Outpatient.xlsx", engine='xlsxwriter') as writer:
                df_mrn.to_excel(writer, index=False, sheet_name="MRN")
                df_merged.to_excel(writer, index=False, sheet_name="New OP")

            with open("Processed_Outpatient.xlsx", "rb") as f:
                st.download_button("Download Processed Workbook", f, file_name="Processed_Outpatient.xlsx")

    elif task == "Sync with SharePoint":
        st.warning("Feature not yet implemented. Requires authentication or OneDrive/SharePoint API.")

    elif task == "Update Notes Column W":
        st.info("Upload both SharePoint and Local workbooks to sync Column W.")
        sharepoint_file = st.file_uploader("Upload SharePoint Excel file:", type="xlsx", key="sharepoint")

        if sharepoint_file:
            xl_local = pd.ExcelFile(uploaded_file)
            xl_sharepoint = pd.ExcelFile(sharepoint_file)
            if "New OP" not in xl_local.sheet_names or "New OP" not in xl_sharepoint.sheet_names:
                st.error("'New OP' sheet missing in one of the files")
            else:
                df_local = xl_local.parse("New OP")
                df_share = xl_sharepoint.parse("New OP")
                df_local["W"] = df_share["W"]

                with pd.ExcelWriter("Updated_Notes.xlsx") as writer:
                    df_local.to_excel(writer, index=False, sheet_name="New OP")

                with open("Updated_Notes.xlsx", "rb") as f:
                    st.download_button("Download Updated Notes", f, file_name="Updated_Notes.xlsx")
else:
    st.info("Please upload an Excel file to begin.")
