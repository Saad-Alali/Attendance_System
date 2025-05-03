import os
import sys
from tkinter import Tk, filedialog

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from project.file_validator import validate_excel_file
from project.browser_automation import automate_browser
from project.excel_processor import merge_excel_files

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
    download_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
    os.makedirs(download_dir, exist_ok=True)
    
    automate_browser(download_dir)
    
    output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "merged_data.xlsx")
    merge_excel_files(download_dir, output_file)
    
    print(f"Data merged successfully in file: {output_file}")

if __name__ == "__main__":
    main()