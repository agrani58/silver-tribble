import pandas as pd
import os

# -------------------------
# Mapping of account numbers to bank abbreviations
account_mapping = {
    "480100000373201": "CBIL",
    "00200105201324": "EBL",
}

# Path to the original Excel file
file_path = r"C:\Users\agran\Downloads\Everest Bank - Statement.xls";

# -------------------------
def process_bank_table(file_path, account_mapping):

    # Read all rows as raw data
    df = pd.read_excel(file_path, header=None)
    df = df.astype(object)

    # -------------------------
    # Find header row containing 'Date' or 'ID'
    header_row_idx = None
    date_col_idx = None
    for col in df.columns:
        matches = df[df[col].astype(str).str.strip().str.contains("Date|ID", case=False, na=False)]
        if not matches.empty:
            header_row_idx = matches.index[0]
            date_col_idx = col
            break
    if header_row_idx is None:
        raise ValueError("Could not find 'Date' or 'ID' in the sheet!")
    
    # -------------------------
    # Detect account number in the table
    account_number_found = None
    for i in range(header_row_idx + 1, len(df)):
        for acc_num in account_mapping:
            if str(acc_num) in str(df.iloc[i].values):
                account_number_found = str(acc_num)
                break
        if account_number_found:
            break
    if account_number_found is None:
        raise ValueError("Account number not found in mapping or table!")
    
    bank_abbr = account_mapping[account_number_found]
    
    # -------------------------
    # Find last valid row in Date column 
    last_table_row_idx = header_row_idx
    for i in range(header_row_idx + 1, len(df)):
        val = df.iat[i, date_col_idx]
        try:
            parsed = pd.to_datetime(val, errors='coerce')
            if not pd.isna(parsed):
                last_table_row_idx = i
        except:
            continue

    # -------------------------
    # Fill each table row individually to avoid gaps
    for i in range(header_row_idx + 1, last_table_row_idx + 1):
        last_col = df.iloc[i, :].last_valid_index()
        df.iat[i, last_col + 1] = account_number_found
        df.iat[i, last_col + 2] = bank_abbr

    # -------------------------
    # Set header names right after the last non-empty column of header
    header_last_col = df.iloc[header_row_idx, :].last_valid_index()
    df.iat[header_row_idx, header_last_col + 1] = "Account Number"
    df.iat[header_row_idx, header_last_col + 2] = "Bank Code"
    
    # -------------------------
    # Grab first valid date for filename
    first_date_str = "nodate"
    for i in range(header_row_idx + 1, last_table_row_idx + 1):
        val = df.iat[i, date_col_idx]
        try:
            parsed = pd.to_datetime(val, errors='coerce')
            if not pd.isna(parsed):
                first_date_str = parsed.strftime("%Y-%m-%d")
                break
        except:
            continue

    # -------------------------

    folder, filename = os.path.split(file_path)
    new_filename = f"{bank_abbr}_statement_{first_date_str}_modified.xlsx"
    new_path = os.path.join(folder, new_filename)


    df.to_excel(new_path, index=False, header=False)

    # Confirmation message
    print(f"Saved modified table as: {new_filename}")
    return new_path

 

process_bank_table(file_path, account_mapping)
