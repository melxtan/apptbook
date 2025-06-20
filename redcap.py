import requests
import pandas as pd

def fetch_redcap_data(api_key: str, redcap_url: str = "https://redcap.med.usc.edu/api/"):
    payload = {
        "token": api_key,
        "content": "record",
        "format": "json",
        "type": "flat",
        "rawOrLabel": "label"
    }
    response = requests.post(redcap_url, data=payload)
    response.raise_for_status()
    return response.json()

def parse_redcap_to_df(json_data: list) -> pd.DataFrame:
    return pd.DataFrame(json_data)

def filter_new_records(df_new: pd.DataFrame, df_existing: pd.DataFrame, case_id_col: str = "full_case_id"):
    existing_ids = df_existing[case_id_col].astype(str).str.strip().unique()
    df_new[case_id_col] = df_new[case_id_col].astype(str).str.strip()
    return df_new[~df_new[case_id_col].isin(existing_ids)]

def update_mrn_sheet(df_existing: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat([df_existing, df_new], ignore_index=True)
    return combined.drop_duplicates(subset="full_case_id")
