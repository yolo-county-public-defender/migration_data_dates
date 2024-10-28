import pandas as pd
import re
import traceback
from openpyxl.styles import Alignment

def is_user_id(value):
    # Check if the value is a 5-digit number
    if pd.isna(value):
        return False
    return bool(re.match(r'^\d{5}$', str(value).strip()))

def clean_and_process_excel():
    try:
        # Read all sheets from the Excel file
        print("Reading Excel file...")
        excel_file = pd.read_excel('migrationdata2.xlsx', sheet_name=None)
        
        # Get the Note sheet
        df = excel_file['Note']
        df_modified = df.copy()
        
        # Step 1: Clean overflow notes and fix user IDs
        print("Cleaning overflow notes and fixing user IDs...")
        for idx in df_modified.index:
            user_id_found = False
            
            # Check columns D through H for user ID and clear non-user ID content
            for col in ['D', 'E', 'F', 'G', 'H']:
                col_idx = ord(col) - ord('A')  # Convert column letter to index
                value = df_modified.iloc[idx, col_idx]
                
                if is_user_id(value):
                    # Found user ID, move it to column E
                    df_modified.iloc[idx, 4] = value  # Column E is index 4
                    if col != 'E':  # If user ID wasn't in column E, clear original location
                        df_modified.iloc[idx, col_idx] = ''
                    user_id_found = True
                elif not pd.isna(value):
                    # Clear non-user ID content
                    df_modified.iloc[idx, col_idx] = ''
        
        # Step 2: Clear "NULL" values in column E
        df_modified.iloc[:, 4] = df_modified.iloc[:, 4].replace('NULL', '')
        
        # Step 3: Process datetime in column B - only keep the date part
        print("Processing datetime values...")
        mask = df_modified.iloc[:, 1].astype(str).str.contains(r'\d{1,2}/\d{1,2}/\d{4}', na=False)
        
        def extract_date(value):
            try:
                dt = pd.to_datetime(value, format='%m/%d/%Y %I:%M%p')
                return dt.strftime('%m/%d/%Y').strip()
            except Exception as e:
                return value.strip() if isinstance(value, str) else value
        
        valid_rows = df_modified[mask]
        for idx, row in valid_rows.iterrows():
            date_time = row.iloc[1]
            try:
                date_only = extract_date(date_time)
                df_modified.iloc[idx, 1] = date_only
            except Exception as e:
                print(f"Error processing row {idx}: {str(e)}")
        
        # Step 4: Update the Note sheet in the original excel_file dictionary
        excel_file['Note'] = df_modified
        
        # Save all sheets to the new Excel file with right alignment for column B
        print("Saving modified Excel file...")
        with pd.ExcelWriter('migrationdatamodified.xlsx', engine='openpyxl') as writer:
            for sheet_name, sheet_df in excel_file.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Only apply right alignment to column B in the Note sheet
                if sheet_name == 'Note':
                    worksheet = writer.sheets[sheet_name]
                    for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):  # Column B
                        for cell in row:
                            cell.alignment = Alignment(horizontal='right')
        
        print("Excel file has been processed successfully!")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("Full error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    clean_and_process_excel()
