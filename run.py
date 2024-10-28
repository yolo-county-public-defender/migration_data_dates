import pandas as pd
import re
import traceback
from openpyxl.styles import Alignment

def is_user_id(value):
    if pd.isna(value):
        return False
    return bool(re.match(r'^\d{5}$', str(value).strip()))

def clean_and_process_excel():
    try:
        print("Reading Excel file...")
        excel_file = pd.read_excel('migrationdata2.xlsx', sheet_name=None)
        
        df = excel_file['Note']
        df_modified = df.copy()
        
        print("Cleaning overflow notes and fixing user IDs...")
        for idx in df_modified.index:
            user_id_found = False
            
            for col in ['D', 'E', 'F', 'G', 'H']:
                col_idx = ord(col) - ord('A')  
                value = df_modified.iloc[idx, col_idx]
                
                if is_user_id(value):
                    df_modified.iloc[idx, 4] = value 
                    if col != 'E':  
                        df_modified.iloc[idx, col_idx] = ''
                    user_id_found = True
                elif not pd.isna(value):
      
                    df_modified.iloc[idx, col_idx] = ''

        df_modified.iloc[:, 4] = df_modified.iloc[:, 4].replace('NULL', '')

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
  
        excel_file['Note'] = df_modified
 
        print("Saving modified Excel file...")
        with pd.ExcelWriter('migrationdatamodified.xlsx', engine='openpyxl') as writer:
            for sheet_name, sheet_df in excel_file.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
 
                if sheet_name == 'Note':
                    worksheet = writer.sheets[sheet_name]
                    for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):  
                        for cell in row:
                            cell.alignment = Alignment(horizontal='right')
        
        print("Excel file has been processed successfully!")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("Full error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    clean_and_process_excel()
