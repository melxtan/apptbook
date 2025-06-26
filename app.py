import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font
from io import BytesIO
from datetime import datetime, timedelta, time

st.set_page_config(layout="wide")
st.title("Automated REDCap & Excel Workflow")

excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

api_keys = [st.secrets["redcap_api_1"], st.secrets["redcap_api_2"]]

# ...[parse_excel_date, parse_to_time, parse_json_to_excel as in your code]...

def move_routine_to_newop(df_newop, df_routine, df_mrn):
    # [Your logic unchanged, but function returns dataframes, not counts]
    df_routine_to_move = df_routine.iloc[5:, :22].copy()
    df_routine_to_move['font_color'] = 'red'
    df_newop = pd.concat([df_newop, df_routine_to_move], ignore_index=True)

    # Formatting columns as in your code
    if 'D' in df_newop.columns:
        df_newop['D'] = pd.to_datetime(df_newop['D'], errors='coerce').dt.strftime('%m/%d/%Y')
    if 'E' in df_newop.columns:
        df_newop['E'] = pd.to_datetime(df_newop['E'], errors='coerce').dt.strftime('%H:%M')
    if 'H' in df_newop.columns:
        df_newop['H'] = pd.to_datetime(df_newop['H'], errors='coerce').dt.strftime('%m/%d/%Y')
    if set(['D', 'E']).issubset(df_newop.columns):
        df_newop = df_newop.sort_values(by=['D', 'E'], ascending=[True, True], na_position='last').reset_index(drop=True)

    # Add columns X, Z, AA if missing
    for col_idx, col_letter in zip([23, 25, 26], ['X', 'Z', 'AA']):
        if len(df_newop.columns) <= col_idx:
            for _ in range(col_idx + 1 - len(df_newop.columns)):
                df_newop[df_newop.shape[1]] = ""

    # VLOOKUP logic
    for idx in range(6, len(df_newop)):
        if pd.isna(df_newop.iloc[idx, 23]) or df_newop.iloc[idx, 23] == "":
            lookup_val = df_newop.iloc[idx, 1]  # Column B
            vlookup_row = df_mrn[df_mrn.iloc[:, 0] == lookup_val]
            if not vlookup_row.empty:
                df_newop.iat[idx, 23] = vlookup_row.iloc[0, 1] if df_mrn.shape[1] > 1 else ""
                df_newop.iat[idx, 25] = vlookup_row.iloc[0, 5] if df_mrn.shape[1] > 5 else ""
                df_newop.iat[idx, 26] = vlookup_row.iloc[0, 6] if df_mrn.shape[1] > 6 else ""

    df_newop = df_newop.drop_duplicates(subset=list(df_newop.columns[:26]), keep='first').reset_index(drop=True)
    df_routine.iloc[5:, :22] = ""

    # Return both updated DataFrames and a marker for highlighting
    return df_newop, df_routine, df_mrn, df_routine_to_move.index + len(df_newop) - len(df_routine_to_move)

def df_to_sheet(ws, df, font_color_indices=None):
    # Overwrites sheet with DataFrame contents.
    for r, row in enumerate(df.itertuples(index=False), 1):
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
            if font_color_indices and (r-1) in font_color_indices:
                ws.cell(row=r, column=c).font = Font(color='FF0000')

if excel_file:
    excel_bytes = BytesIO(excel_file.read())
    wb = load_workbook(excel_bytes)
    sheets = wb.sheetnames

    # Convert sheets to DataFrames
    def sheet_to_df(ws):
        data = list(ws.values)
        return pd.DataFrame(data[1:], columns=data[0]) if data else pd.DataFrame()

    ws_newop = wb["New OP"]
    ws_routine = wb["Routine"]
    ws_mrn = wb["MRN"]
    df_newop = sheet_to_df(ws_newop)
    df_routine = sheet_to_df(ws_routine)
    df_mrn = sheet_to_df(ws_mrn)

    if st.button("Refresh MRN sheet from REDCap (API)"):
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
                # Use openpyxl worksheet for this update
                new_mrn = parse_json_to_excel(data, ws_mrn)
                total_new += new_mrn
        st.success(f"MRN sheet refreshed. {total_new} new MRNs imported.")

    if st.button("Run Outpatient Routine Process"):
        df_newop_updated, df_routine_updated, df_mrn_updated, highlight_indices = move_routine_to_newop(df_newop, df_routine, df_mrn)
        # Write DataFrames back to Excel sheets
        df_to_sheet(ws_newop, df_newop_updated, font_color_indices=highlight_indices)
        df_to_sheet(ws_routine, df_routine_updated)
        df_to_sheet(ws_mrn, df_mrn_updated)
        st.success("Routine → New OP processing done. Newly moved rows are highlighted in red.")

    # Save workbook to BytesIO for download
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
