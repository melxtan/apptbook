import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Automated REDCap & Excel Workflow")

# 1. Upload
csv_file = st.file_uploader("Upload CSV File", type=["csv"])
excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

# 2. Read secrets
api_keys = [st.secrets["redcap_api_1"], st.secrets["redcap_api_2"]]
mrn_pwd = st.secrets["mrn_password"]
op_pwd = st.secrets["op_password"]

def parse_json_to_excel(json_data, ws_mrn):
    # Find all existing full_case_id (I column, which is index 9 in openpyxl, 1-based)
    existing_case_ids = set()
    for row in ws_mrn.iter_rows(min_row=2, max_row=ws_mrn.max_row, min_col=9, max_col=9, values_only=True):
        if row[0]:
            existing_case_ids.add(str(row[0]).strip())
    new_mrn_count = 0
    i = ws_mrn.max_row + 1
    for record in json_data:
        incoming_case_id = str(record.get("full_case_id", "")).strip()
        if record.get("mrn") and incoming_case_id not in existing_case_ids:
            # Write fields, using 1-based column indexing
            fields = [
                "mrn", "case_id", "redcap_event_name", "redcap_repeat_instrument",
                "redcap_repeat_instance", "country_origin", "first_responder",
                "internal_referral", "full_case_id", "arm_label", "email",
                "num_appt", "payer_type", "other_refer", "pt_fn", "pt_ln",
                "pt_dob", "today", "service_line"
            ]
            for col_idx, field in enumerate(fields, 1):
                ws_mrn.cell(row=i, column=col_idx, value=record.get(field, ""))
            new_mrn_count += 1
            i += 1
    return new_mrn_count

def paste_csv_to_routine(ws_routine, csv_df):
    # Paste starting from row 6, col 1 (A6)
    for row_idx, row in enumerate(csv_df.values, 6):
        for col_idx, val in enumerate(row, 1):
            ws_routine.cell(row=row_idx, column=col_idx, value=val)

def move_routine_to_newop(wb):
    ws_routine = wb["Routine"]
    ws_newop = wb["New OP"]
    ws_mrn = wb["MRN"]

    # 1. All text in New OP to black
    for row in ws_newop.iter_rows():
        for cell in row:
            cell.font = Font(color="000000")

    # 2. All Routine A6:V in red
    max_row_routine = ws_routine.max_row
    for row in ws_routine.iter_rows(min_row=6, max_row=max_row_routine, min_col=1, max_col=22): # A-V
        for cell in row:
            cell.font = Font(color="FF0000")

    # 3. Copy A6:V last to New OP last empty row
    values = []
    for row in ws_routine.iter_rows(min_row=6, max_row=ws_routine.max_row, min_col=1, max_col=22, values_only=True):
        if any([cell is not None and str(cell).strip() != "" for cell in row]):
            values.append(row)
    insert_row = ws_newop.max_row + 1
    for r, row in enumerate(values, insert_row):
        for c, val in enumerate(row, 1):
            ws_newop.cell(row=r, column=c, value=val)
            ws_newop.cell(row=r, column=c).font = Font(color="FF0000")

    # 4. Set date & time formats for cols D (4), E (5), H (8)
    for cell in ws_newop[get_column_letter(4)]:
        cell.number_format = 'mm/dd/yyyy'
    for cell in ws_newop[get_column_letter(5)]:
        cell.number_format = 'hh:mm'
    for cell in ws_newop[get_column_letter(8)]:
        cell.number_format = 'mm/dd/yyyy'

    # 5. Sort by D then E
    # This requires a temp DataFrame (assuming header in first row)
    data = []
    for row in ws_newop.iter_rows(values_only=True):
        data.append(list(row))
    df = pd.DataFrame(data)
    if not df.empty and df.shape[0] > 1:
        # sort by D(3) then E(4)
        df_sorted = df.sort_values(by=[3, 4], ascending=[True, True])
        for r_idx, row in enumerate(df_sorted.values.tolist(), 1):
            for c_idx, val in enumerate(row, 1):
                ws_newop.cell(row=r_idx, column=c_idx, value=val)

    # 6. Fill X (24), Z (26), AA (27) using MRN lookups
    # We'll convert MRN sheet to DataFrame for fast lookups
    mrn_data = []
    for row in ws_mrn.iter_rows(values_only=True):
        mrn_data.append(row)
    mrn_df = pd.DataFrame(mrn_data)
    mrn_dict_case = {}
    mrn_dict_country = {}
    mrn_dict_firstres = {}
    if not mrn_df.empty and mrn_df.shape[1] >= 7:
        for idx, row in mrn_df.iterrows():
            if pd.notnull(row[0]): # MRN
                mrn_dict_case[str(row[0])] = row[1] if len(row) > 1 else ""
                mrn_dict_country[str(row[0])] = row[5] if len(row) > 5 else ""
                mrn_dict_firstres[str(row[0])] = row[6] if len(row) > 6 else ""

    # For each cell in column X (24), starting from row 7
    for row_idx in range(7, ws_newop.max_row + 1):
        cell_x = ws_newop.cell(row=row_idx, column=24)
        if cell_x.value in [None, ""]:
            col_b = ws_newop.cell(row=row_idx, column=2).value  # B
            if col_b and str(col_b) in mrn_dict_case:
                cell_x.value = mrn_dict_case.get(str(col_b), "")
                ws_newop.cell(row=row_idx, column=26).value = mrn_dict_country.get(str(col_b), "")
                ws_newop.cell(row=row_idx, column=27).value = mrn_dict_firstres.get(str(col_b), "")

    # 7. Remove duplicates A:AA (1:27), keep first row (header or otherwise)
    df2 = pd.DataFrame([[cell.value for cell in row[:27]] for row in ws_newop.iter_rows()])
    df2 = df2.drop_duplicates()
    for i, row in enumerate(df2.values.tolist(), 1):
        for j, val in enumerate(row, 1):
            ws_newop.cell(row=i, column=j, value=val)
    # Clear any extra rows after dedup
    for row in ws_newop.iter_rows(min_row=len(df2)+1, max_row=ws_newop.max_row):
        for cell in row:
            cell.value = None

    # 8. Clear Routine A6:V*
    for row in ws_routine.iter_rows(min_row=6, max_row=max_row_routine, min_col=1, max_col=22):
        for cell in row:
            cell.value = None

if csv_file and excel_file:
    # Load files
    csv_df = pd.read_csv(csv_file)
    excel_bytes = BytesIO(excel_file.read())
    wb = load_workbook(excel_bytes)

    st.success("Files loaded. Ready to process.")

    if st.button("Paste CSV to Routine tab"):
        ws_routine = wb["Routine"]
        paste_csv_to_routine(ws_routine, csv_df)
        st.success("CSV pasted to Routine sheet (starting at A6).")

    if st.button("Refresh MRN sheet from REDCap (API)"):
        ws_mrn = wb["MRN"]
        total_new = 0
        for key in api_keys:
            response = requests.post(
                "https://redcap.med.usc.edu/api/",
                data={
                    "token": key,
                    "content": "record",
                    "format": "json",
                    "type": "flat",
                    "rawOrLabel": "label"
                }
            )
            if response.ok:
                data = response.json()
                new_mrn = parse_json_to_excel(data, ws_mrn)
                total_new += new_mrn
        st.success(f"MRN sheet refreshed. {total_new} new MRNs imported.")

    if st.button("Run Outpatient Routine Process"):
        move_routine_to_newop(wb)
        st.success("Routine → New OP processing done.")

    output = BytesIO()
    wb.save(output)
    st.download_button(
        label="Download Processed Excel",
        data=output.getvalue(),
        file_name="processed_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.info("Remember: You can now upload this Excel to SharePoint as needed.")

else:
    st.info("Please upload both the CSV and Excel files to begin.")
