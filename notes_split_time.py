import pandas as pd
import traceback
from openpyxl.styles import Alignment
from openpyxl.styles.numbers import FORMAT_DATE_XLSX14, NumberFormat
from datetime import datetime

def clean_and_process_excel():
    try:
        print("Reading Excel file...")
        # Read all sheets
        excel_file = pd.read_excel('migrationdata2.xlsx', sheet_name=None)
        
        # Process only the Note sheet
        df_modified = excel_file['Note'].copy()
        
        print("Processing datetime values...")
        # Convert column B to datetime where possible
        datetime_col = pd.to_datetime(df_modified.iloc[:, 1], errors='coerce')
        
        # Extract date only for column B
        df_modified.iloc[:, 1] = datetime_col.dt.date
        
        # Extract time only for column D
        df_modified.iloc[:, 3] = datetime_col.dt.time
        
        # Update the Note sheet in the dictionary
        excel_file['Note'] = df_modified
        
        print("Saving modified Excel file...")
        with pd.ExcelWriter('migrationdatamodified.xlsx', engine='openpyxl') as writer:
            # Write all sheets
            for sheet_name, df in excel_file.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Only format the Note sheet
                if sheet_name == 'Note':
                    worksheet = writer.sheets['Note']
                    
                    # Format column B (date) cells with openpyxl's date format
                    for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):
                        for cell in row:
                            cell.number_format = FORMAT_DATE_XLSX14
                    
                    # Format column D (time) cells
                    for row in worksheet.iter_rows(min_row=2, min_col=4, max_col=4):
                        for cell in row:
                            cell.number_format = 'hh:mm AM/PM'
        
        print("Excel file has been processed successfully!")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("Full error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    clean_and_process_excel()
