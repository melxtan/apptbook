import pandas as pd
from io import BytesIO
from redcap import fetch_redcap_data, parse_redcap_to_df, filter_new_records, update_mrn_sheet
import streamlit as st


def process_files(csv_file, excel_file):
    df_csv = pd.read_csv(csv_file)
    xl = pd.ExcelFile(excel_file)

    new_op_df = xl.parse("New OP", header=None)
    ip_df = xl.parse("IP", header=None)
    mrn_df = xl.parse("MRN", header=0)  # Use header row for MRN tab

    # Step 1: Update MRN sheet from REDCap
    api_key = st.secrets["REDCAP"]["KEY_1"]
    redcap_data = fetch_redcap_data(api_key)
    df_redcap = parse_redcap_to_df(redcap_data)
    updated_mrn_df = update_mrn_sheet(mrn_df, df_redcap)

    # Step 2: Append CSV to New OP
    new_data = df_csv.copy()
    new_data.reset_index(drop=True, inplace=True)
    combined_new_op = pd.concat([new_op_df, new_data], ignore_index=True)

    # Step 3: Parse columns D, E, H as datetime/time for sorting
    combined_new_op[3] = pd.to_datetime(combined_new_op[3], errors='coerce')  # Column D
    combined_new_op[4] = pd.to_datetime(combined_new_op[4], format="%H:%M", errors='coerce').dt.time  # Column E
    combined_new_op[7] = pd.to_datetime(combined_new_op[7], errors='coerce')  # Column H

    # Step 4: Sort by Date (D) then Time (E)
    combined_new_op = combined_new_op.sort_values(by=[3, 4], ignore_index=True)

    # Step 5: Simulate VLOOKUP
    mrn_lookup = updated_mrn_df.set_index("MRN")
    combined_new_op[23] = combined_new_op[1].map(mrn_lookup["Case ID"].to_dict())
    combined_new_op[25] = combined_new_op[1].map(mrn_lookup.get("Payer Type", pd.Series()).to_dict())
    combined_new_op[26] = combined_new_op[1].map(mrn_lookup.get("Please enter the name of the internal referer", pd.Series()).to_dict())

    # Step 6: Remove duplicates
    combined_new_op = combined_new_op.drop_duplicates(subset=range(0, 27))

    # Step 7: Re-format date/time columns for output
    combined_new_op[3] = combined_new_op[3].dt.strftime("%m/%d/%Y")  # D
    combined_new_op[7] = combined_new_op[7].dt.strftime("%m/%d/%Y")  # H

    # Step 8: Write back to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_new_op.to_excel(writer, index=False, header=False, sheet_name="New OP")
        ip_df.to_excel(writer, index=False, header=False, sheet_name="IP")
        updated_mrn_df.to_excel(writer, index=False, header=True, sheet_name="MRN")

    output.seek(0)
    return output
