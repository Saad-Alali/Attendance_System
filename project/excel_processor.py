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

def split_by_component(merged_file, output_dir):
    try:
        df = pd.read_excel(merged_file)
        
        component_types = ['LEC1', 'LEC2', 'LAB1', 'LAB2']
        
        for component in component_types:
            component_df = df[df['اسم المكون'] == component]
            
            if not component_df.empty:
                output_file = os.path.join(output_dir, f"{component}.xlsx")
                component_df.to_excel(output_file, index=False)
                print(f"Created {component} file with {len(component_df)} records")
    except Exception as e:
        print(f"Error splitting files by component: {str(e)}")