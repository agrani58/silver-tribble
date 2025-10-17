import pandas as pd
import os

def process_bank_table(file_path, account_number, bank_abbr):
    # Read everything as raw data
    df = pd.read_excel(file_path, header=None)
    
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
    
    # Determine actual last column of the table
    last_col_idx = df.iloc[header_row_idx].last_valid_index()
    
    # Insert new columns at the end of the table
    df.insert(last_col_idx + 1, 'Account Number', ['' for _ in range(len(df))])
    df.insert(last_col_idx + 2, 'Bank Code', ['' for _ in range(len(df))])
    
    # Fill values below header
    for i in range(header_row_idx + 1, len(df)):
        df.iat[i, last_col_idx + 1] = account_number
        df.iat[i, last_col_idx + 2] = bank_abbr
    
    # Set header names
    df.iat[header_row_idx, last_col_idx + 1] = 'Account Number'
    df.iat[header_row_idx, last_col_idx + 2] = 'Bank Code'
    
    # Grab first valid date for filename
    first_date_str = "nodate"
    for i in range(header_row_idx + 1, len(df)):
        val = df.iat[i, date_col_idx]
        try:
            parsed = pd.to_datetime(val, errors='coerce')
            if not pd.isna(parsed):
                first_date_str = parsed.strftime("%Y-%m-%d")
                break
        except:
            continue
    
    # Save file
    folder, filename = os.path.split(file_path)
    new_filename = f"{bank_abbr}_statement_{first_date_str}_modified.xlsx"
    new_path = os.path.join(folder, new_filename)
    df.to_excel(new_path, index=False, header=False)
    
    print(f"💾 Saved modified table as: {new_filename}")
    return new_path


bank_files = [
    {"file": r"C:\Users\agran\Downloads\Citizen Bank - Statement.xls", "account": "111111111", "abbr": "CBIL"},
    {"file": r"C:\Users\agran\Downloads\Citizen Bank - Statement.xls", "account": "111111111", "abbr": "CBIL"},
    {"file": r"C:\Users\agran\Downloads\Citizen Bank - Statement.xls", "account": "111111111", "abbr": "CBIL"},
    {"file": r"C:\Users\agran\Downloads\Citizen Bank - Statement.xls", "account": "111111111", "abbr": "CBIL"},
]

for bf in bank_files:
    process_bank_table(bf['file'], bf['account'], bf['abbr'])
