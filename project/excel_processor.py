import os
import pandas as pd
import glob
import numpy as np

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

def update_component_files(original_file, output_dir):
    try:
        original_df = pd.read_excel(original_file)
        
        component_types = ['LEC1', 'LEC2', 'LAB1', 'LAB2']
        student_id_column = 'الرقم الجامعي للطالب'
        
        for component in component_types:
            component_file = os.path.join(output_dir, f"{component}.xlsx")
            
            if not os.path.exists(component_file):
                print(f"Component file {component}.xlsx not found. Skipping.")
                continue
            
            try:
                component_df = pd.read_excel(component_file)
                
                if 'Note' not in component_df.columns:
                    component_df['Note'] = pd.Series(dtype='object')
                
                if 'النتيجة' not in component_df.columns:
                    component_df['النتيجة'] = np.nan
                
                if 'سبب تغيير الدرجة' not in component_df.columns:
                    component_df['سبب تغيير الدرجة'] = pd.Series(dtype='object')
                
                updated_count = 0
                missing_students = []
                
                for _, original_row in original_df.iterrows():
                    student_id = original_row[student_id_column]
                    component_grade = original_row[component]
                    
                    if pd.notna(component_grade) and component_grade != 0:
                        student_in_component = component_df[component_df[student_id_column] == student_id]
                        
                        if not student_in_component.empty:
                            idx = student_in_component.index[0]
                            component_df.at[idx, 'النتيجة'] = component_grade
                            component_df.at[idx, 'سبب تغيير الدرجة'] = 'OE'
                            updated_count += 1
                        else:
                            missing_students.append((student_id, original_row['الاسم الكامل'], component_grade))
                
                if missing_students:
                    print(f"Found {len(missing_students)} students with {component} grades but not in the {component} file")
                    
                    for student_id, student_name, grade in missing_students:
                        new_row = {col: np.nan for col in component_df.columns}
                        new_row[student_id_column] = student_id
                        new_row['الاسم الكامل'] = student_name
                        new_row['النتيجة'] = grade
                        new_row['سبب تغيير الدرجة'] = 'OE'
                        new_row['Note'] = f"Student not originally in {component} list"
                        
                        component_df = pd.concat([component_df, pd.DataFrame([new_row])], ignore_index=True)
                
                component_df.to_excel(component_file, index=False)
                print(f"Updated {component} file: {updated_count} grades updated, {len(missing_students)} students added")
                
            except Exception as e:
                print(f"Error processing {component} file: {str(e)}")
    
    except Exception as e:
        print(f"Error updating component files: {str(e)}")