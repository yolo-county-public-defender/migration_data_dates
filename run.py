import pandas as pd
import re
import traceback

def is_user_id(value):
    # Check if the value is a 5-digit number
    if pd.isna(value):
        return False
    return bool(re.match(r'^\d{5}$', str(value).strip()))

def clean_and_process_excel():
    try:
        # Read all sheets from the Excel file
        print("Reading Excel file...")
        excel_file = pd.read_excel('case_imports_24hr.xlsx', sheet_name=None)
        
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
        
        # Step 3: Process datetime in column B
        print("Processing datetime values...")
        mask = df_modified.iloc[:, 1].astype(str).str.contains(r'\d{1,2}/\d{1,2}/\d{4}', na=False)
        
        def split_datetime(value):
            try:
                dt = pd.to_datetime(value, format='%m/%d/%Y %I:%M%p')
                return pd.Series([dt.strftime('%m/%d/%Y'), dt.strftime('%H:%M')])
            except Exception as e:
                return pd.Series([value, ''])
        
        valid_rows = df_modified[mask]
        for idx, row in valid_rows.iterrows():
            date_time = row.iloc[1]
            try:
                date, time = split_datetime(date_time)
                df_modified.iloc[idx, 1] = date
                df_modified.iloc[idx, 3] = time
            except Exception as e:
                print(f"Error processing row {idx}: {str(e)}")
        
        # Step 4: Update the Note sheet in the original excel_file dictionary
        excel_file['Note'] = df_modified
        
        # Save all sheets to the new Excel file
        print("Saving modified Excel file...")
        with pd.ExcelWriter('case_imports_modified.xlsx') as writer:
            for sheet_name, sheet_df in excel_file.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("Excel file has been processed successfully!")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("Full error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    clean_and_process_excel()
