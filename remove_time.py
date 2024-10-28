import pandas as pd
import traceback

def remove_timestamps():
    try:
        # Read all sheets from the Excel file
        print("Reading Excel file...")
        excel_file = pd.read_excel('migrationdata2.xlsx', sheet_name=None)
        
        # Get the Note sheet
        df = excel_file['Note']
        df_modified = df.copy()
        
        # Process datetime in column B - only keep the date part
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
            date_time = row.iloc[1]  # Column B is index 1
            try:
                date_only = extract_date(date_time)
                df_modified.iloc[idx, 1] = date_only
            except Exception as e:
                print(f"Error processing row {idx}: {str(e)}")
        
        # Update the Note sheet in the original excel_file dictionary
        excel_file['Note'] = df_modified
        
        # Save all sheets to the new Excel file
        print("Saving modified Excel file...")
        with pd.ExcelWriter('migrationdata_times_removed.xlsx') as writer:
            for sheet_name, sheet_df in excel_file.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("Excel file has been processed successfully!")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("Full error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    remove_timestamps()