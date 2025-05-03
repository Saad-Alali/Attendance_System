import os
import pandas as pd
import glob

def merge_excel_files(download_dir, output_file):
    excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    
    if not excel_files:
        print("No Excel files found for merging.")
        return
    
    all_data = []
    
    for file in excel_files:
        try:
            df = pd.read_excel(file)
            
            if len(df) > 1:
                all_data.append(df.iloc[1:])
        except Exception as e:
            print(f"Error reading file {file}: {str(e)}")
    
    if not all_data:
        print("No data to merge.")
        return
    
    try:
        merged_df = pd.concat(all_data, ignore_index=True)
        merged_df.to_excel(output_file, index=False)
    except Exception as e:
        print(f"Error merging files: {str(e)}")