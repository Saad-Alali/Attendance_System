import os
import sys
from tkinter import Tk, filedialog

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from project.file_validator import validate_excel_file
    from project.browser_automation import automate_browser
    from project.excel_processor import merge_excel_files, split_by_component, update_component_files
except ImportError:
    module_path = resource_path('project')
    sys.path.append(module_path)
    from file_validator import validate_excel_file
    from browser_automation import automate_browser
    from excel_processor import merge_excel_files, split_by_component, update_component_files

def main():
    root = Tk()
    root.withdraw()
    
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File (Grade Template)",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    
    if not excel_file_path:
        print("No file selected. Program will close.")
        return
    
    is_valid, message = validate_excel_file(excel_file_path)
    if not is_valid:
        print(f"Error validating file: {message}")
        return
    
    print("File validated successfully. Opening browser...")
    
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "TVTC_Grades")
    os.makedirs(desktop_path, exist_ok=True)
    
    download_dir = os.path.join(desktop_path, "downloads")
    os.makedirs(download_dir, exist_ok=True)
    
    automate_browser(download_dir)
    
    output_file = os.path.join(desktop_path, "merged_data.xlsx")
    merge_excel_files(download_dir, output_file)
    
    print(f"Data merged successfully in file: {output_file}")
    
    split_by_component(output_file, desktop_path)
    print("Files split by component successfully.")
    
    update_component_files(excel_file_path, desktop_path)
    print("Component files updated with grades from original template.")

    deleted_count = 0
    for file in os.listdir(download_dir):
        file_path = os.path.join(download_dir, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
                deleted_count += 1
        except Exception as e:
            print(f"Error deleting file {file_path}: {str(e)}")
    
    print(f"Cleaned up downloads folder: {deleted_count} temporary files deleted.")

if __name__ == "__main__":
    main()