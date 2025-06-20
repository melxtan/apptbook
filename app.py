import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# File upload
st.title("Daily Excel Macro Replacement")
csv_file = st.file_uploader("Upload CSV file", type="csv")
excel_file = st.file_uploader("Upload Excel file", type="xlsx")

if csv_file and excel_file:
    # Load CSV
    df_csv = pd.read_csv(csv_file)

    # Load Excel tabs with no headers
    xl = pd.ExcelFile(excel_file)
    new_op_df = xl.parse("New OP", header=None)
    ip_df = xl.parse("IP", header=None)
    routine_df = xl.parse("Routine", header=4)  # Header on row 5 (0-based index 4)
    mrn_df = xl.parse("MRN", header=None)

    # Step 1: Copy CSV data to Routine (replace existing data below row 5)
    routine_df = df_csv.copy()

    # Step 2: Append Routine to New OP (starting from col A, row 6 in Excel)
    new_op_end = new_op_df.shape[0]
    new_data = routine_df.copy()
    new_data.reset_index(drop=True, inplace=True)
    combined_new_op = pd.concat([new_op_df, new_data], ignore_index=True)

    # Step 3: Format New OP columns (pandas doesn't enforce format but we can ensure dtype)
    def to_datetime_safe(series, fmt):
        try:
            return pd.to_datetime(series, errors='coerce').dt.strftime(fmt)
        except:
            return series

    combined_new_op[3] = to_datetime_safe(combined_new_op[3], "%m/%d/%Y")  # Column D
    combined_new_op[4] = pd.to_datetime(combined_new_op[4], errors='coerce').dt.strftime("%H:%M")  # Column E
    combined_new_op[7] = to_datetime_safe(combined_new_op[7], "%m/%d/%Y")  # Column H

    # Step 4: Sort by D then E
    combined_new_op = combined_new_op.sort_values(by=[3, 4], ignore_index=True)

    # Step 5: VLOOKUP logic based on column 1 (MRN) matching MRN!A
    mrn_lookup = mrn_df.set_index(0)
    combined_new_op[23] = combined_new_op[1].map(mrn_lookup[1].to_dict())  # Column X
    combined_new_op[25] = combined_new_op[1].map(mrn_lookup[5].to_dict())  # Column Z
    combined_new_op[26] = combined_new_op[1].map(mrn_lookup[6].to_dict())  # Column AA

    # Step 6: Remove duplicates on columns A:AA (0-26)
    combined_new_op = combined_new_op.drop_duplicates(subset=range(0, 27))

    # Step 7: Clear Routine tab (keep header)
    cleared_routine_df = pd.DataFrame(columns=routine_df.columns)

    # Export to new Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_new_op.to_excel(writer, index=False, header=False, sheet_name="New OP")
        ip_df.to_excel(writer, index=False, header=False, sheet_name="IP")
        cleared_routine_df.to_excel(writer, index=False, header=True, sheet_name="Routine")
        mrn_df.to_excel(writer, index=False, header=False, sheet_name="MRN")

    st.success("Processing complete. Download updated file below.")
    st.download_button("Download Updated Excel", output.getvalue(), file_name="Processed_Apptbook.xlsx")
