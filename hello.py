import os
import re
import pandas as pd

def normalize_number(value):
    """Normalize numeric strings to plain digit strings."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    
    # Handle scientific notation first
    if 'e' in s.lower():
        try:
            # Convert scientific notation to full number
            s = f"{float(s):.0f}"
        except ValueError:
            pass
    
    # Remove all non-numeric characters
    s = re.sub(r"[^0-9]", "", s)
    return s

def detect_account_number(df, mapping):
    """Try to find an account number from mapping or fallback to any 10–20 digit sequence."""
    # Flatten all values and normalize them
    flat_values = []
    for v in df.values.flatten():
        if pd.notna(v) and str(v).strip():
            normalized = normalize_number(v)
            if 8 <= len(normalized) <= 20:
                flat_values.append(normalized)
    
    # Remove duplicates while preserving order
    seen = set()
    unique_values = []
    for v in flat_values:
        if v not in seen:
            seen.add(v)
            unique_values.append(v)

    # Try mapping match
    for val in unique_values:
        for acc_num, bank_code in mapping.items():
            if normalize_number(acc_num) == val:
                return acc_num, bank_code
            if val in normalize_number(acc_num) or normalize_number(acc_num) in val:
                return acc_num, bank_code

    # Fallback
    all_text = " ".join([str(v) for v in df.values.flatten() if pd.notna(v) and str(v).strip()])
    
    for acc_num, bank_code in mapping.items():
        if acc_num in all_text:
            return acc_num, bank_code
        norm_acc = normalize_number(acc_num)
        if norm_acc in all_text:
            return acc_num, bank_code

    if unique_values:
        return unique_values[0], "UNKNOWN"

    return None, "UNKNOWN"

def is_date(value):
    """Check if a value is a date in common formats"""
    if pd.isna(value) or value == "":
        return False
    date_str = str(value).strip()
    # Check for various date formats
    date_patterns = [
        r'\d{4}-\d{2}-\d{2}',  # YYYY-MM-DD
        r'\d{1,2}/\d{1,2}/\d{4}',  # MM/DD/YYYY or DD/MM/YYYY
        r'\d{1,2}-\d{1,2}-\d{4}',  # MM-DD-YYYY or DD-MM-YYYY
    ]
    for pattern in date_patterns:
        if re.match(pattern, date_str):
            return True
    return False

def find_header_row(df):
    """Find the header row containing transaction column names"""
    header_indicators = [
        r'(?i)date', r'(?i)particulars', r'(?i)description', 
        r'(?i)debit', r'(?i)credit', r'(?i)balance',
        r'(?i)withdraw', r'(?i)deposit', r'(?i)chq', r'(?i)cheque',
        r'(?i)transaction', r'(?i)detail'
    ]
    
    for idx, row in df.iterrows():
        row_text = ' '.join([str(cell).strip() for cell in row if str(cell).strip()])
        matches = sum(1 for pattern in header_indicators if re.search(pattern, row_text, re.IGNORECASE))
        if matches >= 2:  # At least 2 header indicators found
            return idx
    
    return None

def extract_transaction_data(df, file_name):
    """
    Extract transaction data by finding the proper header row and removing opening/closing balances
    """
    # Find header row
    header_row_idx = find_header_row(df)
    
    if header_row_idx is None:
        # Try alternative approach - look for WITHDRAW/DEPOSIT pattern
        for idx, row in df.iterrows():
            row_values = [str(cell).upper().strip() for cell in row if str(cell).strip()]
            if 'WITHDRAW' in row_values and 'DEPOSIT' in row_values:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            return None, "nodate"
    
    # Set columns from header row
    df_clean = df.iloc[header_row_idx + 1:].copy()
    df_clean.columns = [str(col).strip() for col in df.iloc[header_row_idx]]
    
    # Rename TRANSACTION DATE to Date and standardize other column names
    column_renames = {}
    for col in df_clean.columns:
        col_upper = str(col).upper()
        if 'TRANSACTION DATE' in col_upper:
            column_renames[col] = 'Date'
        elif 'VALUE DATE' in col_upper and 'Date' not in df_clean.columns:
            column_renames[col] = 'Date'
        elif 'DESCRIPTION' in col_upper or 'PARTICULARS' in col_upper or 'DETAIL' in col_upper:
            column_renames[col] = 'Description'
        elif 'WITHDRAW' in col_upper:
            column_renames[col] = 'Debit'
        elif 'DEPOSIT' in col_upper:
            column_renames[col] = 'Credit'
    
    if column_renames:
        df_clean = df_clean.rename(columns=column_renames)
    
    # Remove empty rows and columns
    df_clean = df_clean.dropna(how='all')
    df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('^Unnamed', na=False)]
    df_clean = df_clean.loc[:, ~(df_clean.columns.str.strip() == '')]
    
    # Remove ID columns
    id_cols = [col for col in df_clean.columns if re.search(r'(?i)\bid\b', str(col))]
    if id_cols:
        df_clean = df_clean.drop(columns=id_cols)
    
    # Clean data
    df_clean = df_clean.fillna('')
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].astype(str).str.strip()
    
    # NEW: Remove completely empty rows (all columns empty)
    before_empty_removal = len(df_clean)
    df_clean = df_clean[~df_clean.apply(lambda row: all(str(cell).strip() == '' for cell in row), axis=1)]
    after_empty_removal = len(df_clean)
    if before_empty_removal > after_empty_removal:
        print(f"   Removed {before_empty_removal - after_empty_removal} completely empty rows")
    
    # Find date column
    date_columns = [col for col in df_clean.columns if re.search(r'(?i)date', str(col))]
    if not date_columns:
        return df_clean.reset_index(drop=True), "nodate"
    
    date_col = date_columns[0]  # Use the first date column found
    
    # ENHANCED LOGIC: Remove rows with various opening/closing balance patterns
    rows_to_remove = []
    for idx, row in df_clean.iterrows():
        # Check for various opening/closing balance patterns
        row_values = [str(cell).strip() for cell in row]
        
        # Pattern 1: "- - Opening Balance" (3 or more columns with "-" at start)
        if len(df_clean.columns) >= 3:
            # Check first few columns for "-" pattern
            dash_pattern = True
            for i in range(min(3, len(df_clean.columns))):
                if str(row[df_clean.columns[i]]).strip() != "-":
                    dash_pattern = False
                    break
            
            if dash_pattern:
                # Check remaining columns for balance keywords
                for i in range(3, len(df_clean.columns)):
                    cell_value = str(row[df_clean.columns[i]]).lower()
                    if "opening balance" in cell_value or "closing balance" in cell_value:
                        rows_to_remove.append(idx)
                        balance_type = "opening" if "opening" in cell_value else "closing"
                        print(f"   Removed {balance_type} balance row (dash pattern) at index {idx}")
                        break
        
        # Pattern 2: Date + "Opening/Closing Balance" + dashes
        date_value = str(row[date_col]).strip() if date_col in row else ""
        if is_date(date_value):
            # Check if this is a balance row with date in first column and balance in second
            if len(df_clean.columns) >= 2:
                second_col = df_clean.columns[1]
                second_col_value = str(row[second_col]).lower() if second_col in row else ""
                
                # Check for opening/closing balance in second column
                if "opening balance" in second_col_value or "closing balance" in second_col_value:
                    # Check if subsequent columns have dashes
                    dash_count = 0
                    for i in range(2, min(5, len(df_clean.columns))):  # Check next 3 columns
                        if i < len(df_clean.columns):
                            col_value = str(row[df_clean.columns[i]]).strip()
                            if col_value == "-":
                                dash_count += 1
                    
                    # If we have at least 2 dashes in subsequent columns, remove the row
                    if dash_count >= 2 and idx not in rows_to_remove:
                        rows_to_remove.append(idx)
                        balance_type = "opening" if "opening" in second_col_value else "closing"
                        print(f"   Removed {balance_type} balance row (date+balance pattern) at index {idx}")
        
        # Pattern 3: Check description column for balance keywords with dash/empty date
        if 'Description' in df_clean.columns:
            desc_value = str(row['Description']).lower()
            date_value = str(row[date_col]).strip() if date_col in row else ""
            
            # If description has balance keywords and date is dash/empty
            if ("opening balance" in desc_value or "closing balance" in desc_value) and (date_value == "-" or date_value == ""):
                if idx not in rows_to_remove:
                    rows_to_remove.append(idx)
                    balance_type = "opening" if "opening" in desc_value else "closing"
                    print(f"   Removed {balance_type} balance row (description pattern) at index {idx}")
    
    # Remove identified pattern rows
    if rows_to_remove:
        df_clean = df_clean.drop(rows_to_remove)
        df_clean = df_clean.reset_index(drop=True)
    
    # Find first valid date row after pattern removal
    first_date_row_idx = None
    for idx, row in df_clean.iterrows():
        date_val = str(row[date_col]).strip()
        if is_date(date_val):
            first_date_row_idx = idx
            break
    
    # Check row above first date row for opening balance (additional safety)
    if first_date_row_idx is not None and first_date_row_idx > 0 and first_date_row_idx < len(df_clean):
        row_above = df_clean.iloc[first_date_row_idx - 1]
        row_above_date = str(row_above[date_col]).strip()
        row_above_desc = str(row_above.get('Description', '')).lower() if 'Description' in row_above else ""
        
        # Check if row above has "-" or empty date and contains balance keywords
        if (row_above_date == "-" or row_above_date == "") and ("opening balance" in row_above_desc or "closing balance" in row_above_desc):
            if (first_date_row_idx - 1) not in rows_to_remove:
                df_clean = df_clean.drop(first_date_row_idx - 1)
                balance_type = "opening" if "opening" in row_above_desc else "closing"
                print(f"   Removed {balance_type} balance row above first date row")
                df_clean = df_clean.reset_index(drop=True)
                # Re-find first date row after deletion
                first_date_row_idx = None
                for idx, row in df_clean.iterrows():
                    date_val = str(row[date_col]).strip()
                    if is_date(date_val):
                        first_date_row_idx = idx
                        break
    
    # Find all date row indices
    date_row_indices = []
    for idx, row in df_clean.iterrows():
        date_val = str(row[date_col]).strip()
        if is_date(date_val):
            date_row_indices.append(idx)
    
    # Check row below second last date row for closing balance (additional safety)
    if len(date_row_indices) >= 2:
        second_last_date_row_idx = date_row_indices[-2]
        
        if second_last_date_row_idx < len(df_clean) - 1:
            row_below = df_clean.iloc[second_last_date_row_idx + 1]
            row_below_date = str(row_below[date_col]).strip()
            row_below_desc = str(row_below.get('Description', '')).lower() if 'Description' in row_below else ""
            
            # Check if row below has "-" or empty date and contains closing balance
            if (row_below_date == "-" or row_below_date == "") and "closing balance" in row_below_desc:
                if (second_last_date_row_idx + 1) not in rows_to_remove:
                    df_clean = df_clean.drop(second_last_date_row_idx + 1)
                    print(f"   Removed closing balance row below second last date row")
    
    # Reset index after potential deletion
    df_clean = df_clean.reset_index(drop=True)
    
    # Recalculate date row indices after all deletions
    date_row_indices = []
    for idx, row in df_clean.iterrows():
        date_val = str(row[date_col]).strip()
        if is_date(date_val):
            date_row_indices.append(idx)
    
    # Remove footer rows (everything after the last date row)
    if date_row_indices:
        last_date_row_idx = date_row_indices[-1]
        
        # Remove all rows after the last date row
        if last_date_row_idx < len(df_clean) - 1:
            rows_removed = len(df_clean) - (last_date_row_idx + 1)
            df_clean = df_clean.iloc[:last_date_row_idx + 1]
            print(f"   Removed {rows_removed} footer rows")
    
    # NEW: Remove any remaining completely empty rows at the beginning
    while len(df_clean) > 0 and all(str(cell).strip() == '' for cell in df_clean.iloc[0]):
        df_clean = df_clean.iloc[1:]
        print(f"   Removed empty row at the beginning")
        df_clean = df_clean.reset_index(drop=True)
    
    # Extract date for filename - clean datetime format
    date_str = "nodate"
    if date_row_indices:
        # Use the first date found in the transactions
        for idx in date_row_indices:
            date_val = str(df_clean.iloc[idx][date_col]).strip()
            # Try to extract date in YYYY-MM-DD format
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', date_val)
            if date_match:
                date_str = date_match.group(1)
                break
            # If no YYYY-MM-DD format, try MM/DD/YYYY
            date_match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_val)
            if date_match:
                month, day, year = date_match.groups()
                date_str = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                break
    
    return df_clean.reset_index(drop=True), date_str, date_row_indices

def clean_bank_statement(file_path, mapping, output_dir):
    try:
        file_name = os.path.basename(file_path)
        print(f"📂 Processing: {file_name}")
        
        # Read Excel file
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
            else:
                df = pd.read_excel(file_path, header=None, dtype=str, engine='openpyxl')
        except Exception as e:
            print(f"   Failed to read file: {e}")
            return None
        
        # Clean the data
        df = df.map(lambda x: str(x).strip() if pd.notna(x) else "")
        
        # Detect account number
        acc_num, bank_code = detect_account_number(df, mapping)

        if bank_code == "UNKNOWN":
            print(f"   Account number not found")
            return None

        # Extract transaction data with error handling
        try:
            df_clean, date_str, date_row_indices = extract_transaction_data(df, file_name)
        except Exception as e:
            print(f"   Error in extract_transaction_data: {e}")
            return None

        if df_clean is None or len(df_clean) == 0:
            print(f"   No transaction data found")
            return None

        # Add metadata columns ONLY from first date row onward
        if date_row_indices:
            first_date_row_idx = date_row_indices[0]
            
            # Initialize Account Number and Bank Code columns with empty strings
            df_clean["Account Number"] = ""
            df_clean["Bank Code"] = ""
            
            # Fill Account Number and Bank Code only from first date row to the end
            df_clean.loc[first_date_row_idx:, "Account Number"] = acc_num
            df_clean.loc[first_date_row_idx:, "Bank Code"] = bank_code
            print(f"   Added Account Number and Bank Code from row {first_date_row_idx + 1} onward")
        else:
            # Fallback: add to all rows if no dates found
            df_clean["Account Number"] = acc_num
            df_clean["Bank Code"] = bank_code
            print(f"   Added Account Number and Bank Code to all rows (no dates found)")

        # Create output directory
        os.makedirs(output_dir, exist_ok=True)

        # Construct output path - ensure valid filename
        output_filename = f"{bank_code}_statement_{date_str}_cleaned.xlsx"
        output_path = os.path.abspath(os.path.join(output_dir, output_filename))

        # Save file
        df_clean.to_excel(output_path, index=False, engine="openpyxl")

        print(f"✅ Saved: {output_filename}")
        print(f"   Columns: {list(df_clean.columns)}")
        print(f"   Transaction Rows: {len(df_clean)}")

        return output_path

    except Exception as e:
        print(f"   Error: {e}")
        return None


if __name__ == "__main__":
    base_dir = os.getcwd()
    input_dir = r"C:\Users\agran\OneDrive\Desktop\statements"
    output_dir = os.path.join(base_dir, "output")

    # Your account number mapping
    account_mapping = {
        "480100000373201": "CBIL",
        "00200105201324": "EBL",
        "04000100783692000002": "GBB",
        "00100100149847000001": "JBBL",
        "0050019841800001": "KBL",
        "06211001182": "LAMI",
        "0390063440500013": "MBL",
    }

    processed, errors = 0, 0

    # Process ALL Excel files
    for file_name in os.listdir(input_dir):
        if file_name.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(input_dir, file_name)
            result = clean_bank_statement(file_path, account_mapping, output_dir)
            if result:
                processed += 1
            else:
                errors += 1
                print(f"❌ Skipped: {file_name}")

    print(f"\n📊 Complete! Processed: {processed}, Errors: {errors}")