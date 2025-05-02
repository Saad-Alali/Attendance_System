import time
import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.edge.options import Options as EdgeOptions
import os

INITIAL_URL = "https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx"
TARGET_URL = "https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/facultyAttendanceTracking#!/markAttendance"

pyautogui.PAUSE = 1.0
pyautogui.FAILSAFE = True

def prompt_for_excel_file():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        print("No file selected. Exiting program.")
        exit()
    
    return file_path

def select_browser():
    root = tk.Tk()
    root.title("Select Browser")
    
    browser_var = tk.StringVar(value="chrome")
    
    tk.Label(root, text="Select browser to use:").pack(pady=10)
    
    tk.Radiobutton(root, text="Chrome", variable=browser_var, value="chrome").pack(anchor=tk.W)
    tk.Radiobutton(root, text="Edge", variable=browser_var, value="edge").pack(anchor=tk.W)
    
    selected_browser = [browser_var.get()]
    
    def on_ok():
        selected_browser[0] = browser_var.get()
        root.destroy()
    
    tk.Button(root, text="OK", command=on_ok).pack(pady=10)
    
    root.geometry("300x150")
    root.mainloop()
    
    return selected_browser[0]

def setup_browser(browser_type="chrome"):
    if browser_type == "chrome":
        options = Options()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_experimental_option("detach", True)
        
        print("Starting Chrome browser...")
        driver = webdriver.Chrome(options=options)
    else:
        options = EdgeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_experimental_option("detach", True)
        
        print("Starting Edge browser...")
        driver = webdriver.Edge(options=options)
    
    driver.maximize_window()
    driver.get(INITIAL_URL)
    
    return driver

def find_and_click_image(image_path, max_attempts=15, timeout=3):
    print(f"Looking for image: {os.path.basename(image_path)}")
    for attempt in range(max_attempts):
        try:
            # زيادة نسبة الثقة للبحث عن الصورة
            location = pyautogui.locateCenterOnScreen(image_path, confidence=0.7)
            if location:
                x, y = location
                # جعل حركة الماوس أبطأ (مدة 1.5 ثانية للوصول للهدف)
                pyautogui.moveTo(x, y, duration=1.5)
                time.sleep(1)
                pyautogui.click(x, y)
                print(f"Successfully clicked on {os.path.basename(image_path)}")
                return True
            time.sleep(timeout)
        except Exception as e:
            time.sleep(timeout)
    
    print(f"Failed to find image: {os.path.basename(image_path)}")
    return False

def upload_file(file_path):
    time.sleep(3)
    
    # حل مشكلة التعامل مع نافذة اختيار الملفات
    # نستخدم Ctrl+L للانتقال إلى شريط العنوان
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(1)
    
    # مسح أي نص موجود
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    
    # كتابة مسار الملف كاملاً
    pyautogui.write(file_path)
    time.sleep(1)
    
    # الضغط على Enter للتأكيد
    pyautogui.press('enter')
    time.sleep(3)

def monitor_url_and_automate(driver, excel_file_path, image_paths):
    print("Waiting for new tab to open...")
    initial_handle = driver.current_window_handle
    initial_handles_count = len(driver.window_handles)
    
    while True:
        current_handles = driver.window_handles
        
        if len(current_handles) > initial_handles_count:
            new_handles = [h for h in current_handles if h != initial_handle]
            if new_handles:
                new_tab = new_handles[-1]
                driver.switch_to.window(new_tab)
                print("New tab detected. Switched to new tab.")
                break
        
        time.sleep(1)
    
    print("Monitoring URL in new tab...")
    
    while True:
        current_url = driver.current_url
        
        if current_url == TARGET_URL or current_url.endswith("markAttendance"):
            print("Target URL detected in new tab")
            
            time.sleep(5)
            
            for img_path in image_paths[:2]:
                if find_and_click_image(img_path):
                    time.sleep(3)
            
            print(f"Uploading Excel file: {excel_file_path}")
            upload_file(excel_file_path)
            
            for img_path in image_paths[2:]:
                if find_and_click_image(img_path):
                    time.sleep(3)
            
            print("Automation completed")
            break
        
        time.sleep(1)

def main():
    print("\n=== TVTC Automation Tool ===\n")
    
    browser_type = select_browser()
    excel_file_path = prompt_for_excel_file()
    print(f"Selected Excel file: {excel_file_path}")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_paths = [
        os.path.join(script_dir, "image.png"),
        os.path.join(script_dir, "image1.png"),
        os.path.join(script_dir, "image2.png"),
        os.path.join(script_dir, "image3.png"),
        os.path.join(script_dir, "image4.png"),
        os.path.join(script_dir, "image5.png"),
    ]
    
    missing_images = [img_path for img_path in image_paths if not os.path.exists(img_path)]
    if missing_images:
        print("Warning: The following image files were not found:")
        for img_path in missing_images:
            print(f"  - {os.path.basename(img_path)}")
        
        confirm = messagebox.askyesno(
            "Missing Images",
            "Some image files are missing. Continue anyway?"
        )
        
        if not confirm:
            print("Operation cancelled by user.")
            return
    
    try:
        driver = setup_browser(browser_type)
        monitor_url_and_automate(driver, excel_file_path, image_paths)
        
        print("\nAutomation complete. Browser will remain open.")
        
    except Exception as e:
        print(f"\n❌ An error occurred: {str(e)}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    print("\n=== Script Finished ===")

if __name__ == "__main__":
    main()