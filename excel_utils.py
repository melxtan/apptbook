import pandas as pd
from io import BytesIO

def process_files(csv_file, excel_file):
    df_csv = pd.read_csv(csv_file)
    xl = pd.ExcelFile(excel_file)

    new_op_df = xl.parse("New OP", header=None)
    ip_df = xl.parse("IP", header=None)
    mrn_df = xl.parse("MRN", header=None)

    # Step 1: Append CSV to New OP
    new_data = df_csv.copy()
    new_data.reset_index(drop=True, inplace=True)
    combined_new_op = pd.concat([new_op_df, new_data], ignore_index=True)

    # Step 2: Parse columns D, E, H as datetime/time for sorting
    combined_new_op[3] = pd.to_datetime(combined_new_op[3], errors='coerce')  # Column D
    combined_new_op[4] = pd.to_datetime(combined_new_op[4], format="%H:%M", errors='coerce').dt.time  # Column E
    combined_new_op[7] = pd.to_datetime(combined_new_op[7], errors='coerce')  # Column H

    # Step 3: Sort by Date (D) then Time (E)
    combined_new_op = combined_new_op.sort_values(by=[3, 4], ignore_index=True)

    # Step 4: Simulate VLOOKUP
    mrn_lookup = mrn_df.set_index(0)
    combined_new_op[23] = combined_new_op[1].map(mrn_lookup[1].to_dict())
    combined_new_op[25] = combined_new_op[1].map(mrn_lookup[5].to_dict())
    combined_new_op[26] = combined_new_op[1].map(mrn_lookup[6].to_dict())

    # Step 5: Remove duplicates
    combined_new_op = combined_new_op.drop_duplicates(subset=range(0, 27))

    # Step 6: Re-format date/time columns for output
    combined_new_op[3] = combined_new_op[3].dt.strftime("%m/%d/%Y")  # D
    combined_new_op[7] = combined_new_op[7].dt.strftime("%m/%d/%Y")  # H

    # Step 7: Write back to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_new_op.to_excel(writer, index=False, header=False, sheet_name="New OP")
        ip_df.to_excel(writer, index=False, header=False, sheet_name="IP")
        mrn_df.to_excel(writer, index=False, header=False, sheet_name="MRN")

    output.seek(0)
    return output
