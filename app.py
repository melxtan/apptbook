import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font
from io import BytesIO
from dateutil import parser

st.set_page_config(layout="wide")
st.title("Automated REDCap & Excel Workflow")

excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

api_keys = [st.secrets["redcap_api_1"], st.secrets["redcap_api_2"]]

def parse_excel_date(val):
    """Try to parse any value to a date, else return original."""
    if val is None or val == "":
        return None
    if isinstance(val, (pd.Timestamp, )):
        return val.to_pydatetime()
    if isinstance(val, str):
        try:
            dt = parser.parse(val)
            return dt
        except Exception:
            return val
    return val

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

def move_routine_to_newop(wb):
    ws_routine = wb["Routine"]
    ws_newop = wb["New OP"]
    ws_mrn = wb["MRN"]

    # 1. Get all Routine data from A6:V*
    data_rows = []
    mrn_set = set()
    for i in range(6, ws_routine.max_row + 1):
        row_vals = [ws_routine.cell(row=i, column=c).value for c in range(1, 23)]
        if any([v is not None and str(v).strip() != "" for v in row_vals]):
            data_rows.append(row_vals)
            if row_vals[0]:
                mrn_set.add(str(row_vals[0]))

    if not data_rows:
        return 0

    # 2. Copy new rows to New OP at the bottom, D and H as parsed dates
    insert_row = ws_newop.max_row + 1
    for r, row in enumerate(data_rows, insert_row):
        for c, val in enumerate(row, 1):
            if c in [4, 8]:  # D or H
                cell_value = parse_excel_date(val)
            else:
                cell_value = val
            ws_newop.cell(row=r, column=c, value=cell_value)

    # 3. Set date & time formats for cols D (4), E (5), H (8)
    for cell in ws_newop['D']:
        cell.number_format = 'mm/dd/yyyy'
    for cell in ws_newop['E']:
        cell.number_format = 'hh:mm'
    for cell in ws_newop['H']:
        cell.number_format = 'mm/dd/yyyy'

    # 4. Deduplicate: treat first row as header, dedup A:AA (columns 1-27)
    all_data = []
    for row in ws_newop.iter_rows(values_only=True):
        all_data.append(list(row) + [None]*(27 - len(row)))
    if len(all_data) < 2:
        return 0

    header = all_data[0]
    data = all_data[1:]
    df = pd.DataFrame(data)
    df_dedup = df.drop_duplicates(subset=list(range(27)), keep='first')
    ws_newop.delete_rows(2, ws_newop.max_row - 1)
    for i, row in enumerate(df_dedup.values.tolist(), 2):
        for j, val in enumerate(row, 1):
            # Convert to date for D/H
            if j in [4, 8]:
                cell_value = parse_excel_date(val)
            else:
                cell_value = val
            ws_newop.cell(row=i, column=j, value=cell_value)

    # 5. VLOOKUPs: Fill X (24), Z (26), AA (27) using MRN lookups
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

    # 6. Mark in red only the new rows (from current Routine), that survived dedup
    red_count = 0
    for row in ws_newop.iter_rows(min_row=2, max_row=ws_newop.max_row):
        if row[0].value and str(row[0].value) in mrn_set:
            for cell in row[:27]:
                cell.font = Font(color="FF0000")
            red_count += 1

    # 7. Clear Routine A6:V*
    for i in range(6, ws_routine.max_row + 1):
        for c in range(1, 23):
            ws_routine.cell(row=i, column=c).value = None

    return red_count

if excel_file:
    excel_bytes = BytesIO(excel_file.read())
    wb = load_workbook(excel_bytes)

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
        red_count = move_routine_to_newop(wb)
        st.success(f"Routine → New OP processing done. {red_count} newly unique records highlighted in red.")

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
    st.info("Please upload your Excel file (with CSV data already pasted into Routine tab, A6).")
