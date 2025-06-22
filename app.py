import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(layout="wide")
st.title("Automated REDCap & Excel Workflow")

excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

api_keys = [st.secrets["redcap_api_1"], st.secrets["redcap_api_2"]]

def parse_excel_date(val):
    """Convert Excel serial or string into datetime or return original if not parseable."""
    if val is None or val == "":
        return None

    # Excel date serial number (assuming 1900 date system)
    if isinstance(val, (int, float)):
        # Excel's day 1 is 1899-12-31, but due to leap year bug, we subtract 2
        return datetime(1899, 12, 30) + timedelta(days=float(val))

    # If it's already a datetime object
    if isinstance(val, datetime):
        return val

    # Try parsing common date/time string formats
    str_val = str(val).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%Y/%m/%d", "%m/%d/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str_val, fmt)
        except ValueError:
            continue

    # Fallback: try pandas parser if available (optional)
    try:
        import pandas as pd
        return pd.to_datetime(str_val, errors='coerce')
    except:
        pass

    # If all parsing fails, return original
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

    # [Unprotect "New OP"] - Cannot actually enforce with openpyxl, so comment for parity
    # ws_newop.protection.sheet = False  # only works in GUI

    # 1. Change all text in "New OP" to black
    for row in ws_newop.iter_rows():
        for cell in row:
            cell.font = Font(color="000000")

    # 2. Get all Routine data from A6:V*
    data_rows = []
    for i in range(6, ws_routine.max_row + 1):
        row_vals = [ws_routine.cell(row=i, column=c).value for c in range(1, 23)]
        if any(v not in [None, ""] for v in row_vals):
            data_rows.append(row_vals)

    if not data_rows:
        print("No new data found in 'Routine'.")
        return 0

    # 3. Copy new rows to "New OP" and format
    insert_row = ws_newop.max_row + 1
    copied_mrns = set()
    for r_idx, row in enumerate(data_rows, insert_row):
        for c, val in enumerate(row, 1):
            cell_value = parse_excel_date(val) if c in [4, 8] else val
            cell = ws_newop.cell(row=r_idx, column=c, value=cell_value)
            cell.font = Font(color="FF0000")  # red font for new rows
        if row[0]:
            copied_mrns.add(str(row[0]))

    # 4. Set formats for columns D, E, H
    for col_letter in ["D", "E", "H"]:
        for cell in ws_newop[col_letter]:
            if col_letter == "D" or col_letter == "H":
                cell.number_format = "mm/dd/yyyy"
            elif col_letter == "E":
                cell.number_format = "hh:mm"

    # 5. Deduplicate on A:AA (columns 1–27)
    all_data = []
    for row in ws_newop.iter_rows(values_only=True):
        all_data.append(list(row) + [None]*(27 - len(row)))

    header = all_data[0]
    data = all_data[1:]
    df = pd.DataFrame(data)
    df_dedup = df.drop_duplicates(subset=list(range(26)), keep='first')

    ws_newop.delete_rows(2, ws_newop.max_row - 1)
    for i, row in enumerate(df_dedup.values.tolist(), 2):
        for j, val in enumerate(row, 1):
            cell = ws_newop.cell(row=i, column=j)
            cell.value = parse_excel_date(val) if j in [4, 8] else val
            # Reset font to black first
            cell.font = Font(color="000000")
            # Re-apply red if MRN is in copied set
            if j == 1 and val and str(val) in copied_mrns:
                for c in range(1, 28):
                    ws_newop.cell(row=i, column=c).font = Font(color="FF0000")

    # 6. Sort by D then E
    df_sorted = df_dedup.sort_values(by=[3, 4])  # Columns D and E are index 3, 4 (0-based)
    ws_newop.delete_rows(2, ws_newop.max_row - 1)
    for i, row in enumerate(df_sorted.values.tolist(), 2):
        for j, val in enumerate(row, 1):
            cell = ws_newop.cell(row=i, column=j)
            cell.value = parse_excel_date(val) if j in [4, 8] else val

    # 7. VLOOKUP-like fill for X, Z, AA (cols 24, 26, 27)
    mrn_data = pd.DataFrame(ws_mrn.values)
    mrn_dict_case = {}
    mrn_dict_country = {}
    mrn_dict_firstres = {}

    for idx, row in mrn_data.iterrows():
        if pd.notnull(row[0]):
            mrn = str(row[0])
            mrn_dict_case[mrn] = row[1] if len(row) > 1 else ""
            mrn_dict_country[mrn] = row[5] if len(row) > 5 else ""
            mrn_dict_firstres[mrn] = row[6] if len(row) > 6 else ""

    for row_idx in range(7, ws_newop.max_row + 1):
        col_b = ws_newop.cell(row=row_idx, column=2).value
        if col_b:
            key = str(col_b)
            if key in mrn_dict_case:
                ws_newop.cell(row=row_idx, column=24).value = mrn_dict_case[key]
                ws_newop.cell(row=row_idx, column=26).value = mrn_dict_country.get(key, "")
                ws_newop.cell(row=row_idx, column=27).value = mrn_dict_firstres.get(key, "")

    # 8. Clear Routine!A6:V*
    for i in range(6, ws_routine.max_row + 1):
        for c in range(1, 23):
            ws_routine.cell(row=i, column=c).value = None

    # [Re-protect "New OP"] - simulation only
    # ws_newop.protection.sheet = True

    print("Process completed successfully.")
    return len(data_rows)

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
