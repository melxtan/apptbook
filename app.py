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
    existing_case_ids = set()
    for row in ws_mrn.iter_rows(min_row=2, max_row=ws_mrn.max_row, min_col=9, max_col=9, values_only=True):
        if row[0]:
            existing_case_ids.add(str(row[0]).strip())
    new_mrn_count = 0
    i = ws_mrn.max_row + 1
    for record in json_data:
        incoming_case_id = str(record.get("full_case_id", "")).strip()
        if record.get("mrn") and incoming_case_id not in existing_case_ids:
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

    # 2. All Routine A6:V in red (visual only, not moved yet)
    max_row_routine = ws_routine.max_row
    for row in ws_routine.iter_rows(min_row=6, max_row=max_row_routine, min_col=1, max_col=22):
        for cell in row:
            cell.font = Font(color="FF0000")

    # 3. Copy A6:V last to New OP last empty row, and mark those as "NEW" in marker_col
    values = []
    for row in ws_routine.iter_rows(min_row=6, max_row=ws_routine.max_row, min_col=1, max_col=22, values_only=True):
        if any([cell is not None and str(cell).strip() != "" for cell in row]):
            values.append(row)
    insert_row = ws_newop.max_row + 1
    marker_col = 28  # AB column, to track newly added rows
    for r, row in enumerate(values, insert_row):
        for c, val in enumerate(row, 1):
            ws_newop.cell(row=r, column=c, value=val)
        ws_newop.cell(row=r, column=marker_col, value="NEW")  # set marker for red

    # 4. Set date & time formats for cols D (4), E (5), H (8)
    for cell in ws_newop[get_column_letter(4)]:
        cell.number_format = 'mm/dd/yyyy'
    for cell in ws_newop[get_column_letter(5)]:
        cell.number_format = 'hh:mm'
    for cell in ws_newop[get_column_letter(8)]:
        cell.number_format = 'mm/dd/yyyy'

    # 5. Read sheet to DataFrame including marker column
    data = []
    for row in ws_newop.iter_rows(values_only=True):
        data.append(list(row[:marker_col]))  # includes marker col
    df = pd.DataFrame(data)
    # sort by D(3) then E(4) if enough columns
    if not df.empty and df.shape[0] > 1 and df.shape[1] >= marker_col:
        df_sorted = df.sort_values(by=[3, 4], ascending=[True, True])
        # Remove duplicates on columns A:AA (0:26)
        df_sorted = df_sorted.drop_duplicates(subset=list(range(27)))
        # Write back to sheet
        for r_idx, row in enumerate(df_sorted.values.tolist(), 1):
            for c_idx, val in enumerate(row, 1):
                ws_newop.cell(row=r_idx, column=c_idx, value=val)
        # Clear extra rows after dedup
        for row in ws_newop.iter_rows(min_row=len(df_sorted)+1, max_row=ws_newop.max_row):
            for cell in row:
                cell.value = None

    # 6. VLOOKUPs: Fill X (24), Z (26), AA (27) using MRN lookups
    mrn_data = []
    for row in ws_mrn.iter_rows(values_only=True):
        mrn_data.append(row)
    mrn_df = pd.DataFrame(mrn_data)
    mrn_dict_case = {}
    mrn_dict_country = {}
    mrn_dict_firstres = {}
    if not mrn_df.empty and mrn_df.shape[1] >= 7:
        for idx, row in mrn_df.iterrows():
            if pd.notnull(row[0]):
                mrn_dict_case[str(row[0])] = row[1] if len(row) > 1 else ""
                mrn_dict_country[str(row[0])] = row[5] if len(row) > 5 else ""
                mrn_dict_firstres[str(row[0])] = row[6] if len(row) > 6 else ""
    for row_idx in range(7, ws_newop.max_row + 1):
        cell_x = ws_newop.cell(row=row_idx, column=24)
        if cell_x.value in [None, ""]:
            col_b = ws_newop.cell(row=row_idx, column=2).value  # B
            if col_b and str(col_b) in mrn_dict_case:
                cell_x.value = mrn_dict_case.get(str(col_b), "")
                ws_newop.cell(row=row_idx, column=26).value = mrn_dict_country.get(str(col_b), "")
                ws_newop.cell(row=row_idx, column=27).value = mrn_dict_firstres.get(str(col_b), "")

    # 7. Re-apply red font for rows where marker is set
    for row in ws_newop.iter_rows(min_row=1, max_row=ws_newop.max_row):
        if row[marker_col - 1].value == "NEW":
            for cell in row[:27]:
                cell.font = Font(color="FF0000")
            row[marker_col - 1].value = None  # clear marker

    # 8. Clear Routine A6:V*
    for row in ws_routine.iter_rows(min_row=6, max_row=max_row_routine, min_col=1, max_col=22):
        for cell in row:
            cell.value = None

if csv_file and excel_file:
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
        st.success("Routine → New OP processing done. Newly moved rows are highlighted in red.")

    output = BytesIO()
    wb.save(output)
    st.download_button(
        label="Download Processed Excel",
        data=output.getvalue(),
        file_name="processed_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.info("You can now upload this Excel to SharePoint if needed.")

else:
    st.info("Please upload both the CSV and Excel files to begin.")
