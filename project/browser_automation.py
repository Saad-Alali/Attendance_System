import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from project.utils import wait_for_download

def automate_browser(download_dir):
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    driver = webdriver.Chrome(options=chrome_options)
    
    driver.get("https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx")
    
    print("Please login manually...")
    
    target_tab = None
    main_window = driver.current_window_handle
    
    while True:
        time.sleep(1)
        current_handles = driver.window_handles
        
        for handle in current_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                current_url = driver.current_url
                if "rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry" in current_url and "#/gradebook" in current_url:
                    target_tab = handle
                    break
        
        if target_tab:
            driver.switch_to.window(target_tab)
            print("Grade page found!")
            break
            
    print("Accessing grade page. Starting process...")
    
    process_subjects(driver, download_dir)
    
    driver.quit()

def process_subjects(driver, download_dir):
    wait = WebDriverWait(driver, 10)
    
    try:
        time.sleep(1)
        print("Searching for subjects...")
        
        subject_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[@data-name='subject']")))
        num_subjects = len(subject_rows)
        print(f"Found {num_subjects} subjects.")
        
        for i in range(num_subjects):
            print(f"Processing subject {i+1} of {num_subjects}...")
            subject_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[@data-name='subject']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", subject_rows[i])
            time.sleep(0.5)
            subject_rows[i].click()
            
            time.sleep(1)
            
            try:
                components_button = wait.until(EC.element_to_be_clickable((By.ID, "componentsButton")))
                components_button.click()
                
                time.sleep(1)
                
                print(f"Starting to process components for subject {i+1}...")
                process_components(driver, wait, download_dir)
                
                print(f"Returning to subjects list...")
                try:
                    gradebook_tab = wait.until(EC.element_to_be_clickable((By.ID, "gradebook-tab")))
                    gradebook_tab.click()
                except (TimeoutException, NoSuchElementException):
                    driver.execute_script("window.history.go(-1)")
                
                time.sleep(1)
            except Exception as e:
                print(f"Error with components button, trying to return to subjects: {str(e)}")
                driver.get("https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry#/gradebook")
                time.sleep(2)
    except Exception as e:
        print(f"Error processing subjects: {str(e)}")

def process_components(driver, wait, download_dir):
    try:
        time.sleep(2)
        print("Searching for components...")
        
        component_count = 0
        try:
            rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[td]")))
            component_count = len(rows)
            print(f"Found {component_count} components.")
        except Exception:
            print("No components found with first selector, trying alternative...")
            try:
                rows = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
                component_count = len([row for row in rows if row.find_elements(By.TAG_NAME, "td")])
                print(f"Found {component_count} components with alternative method.")
            except Exception as e:
                print(f"Could not determine component count: {str(e)}")
                component_count = 4
        
        if component_count == 0:
            print("No components found. Taking screenshot for debugging...")
            screenshot_path = os.path.join(download_dir, "debug_screenshot.png")
            driver.save_screenshot(screenshot_path)
            return
        
        for i in range(component_count):
            try:
                print(f"Processing component {i+1} of {component_count}...")
                
                try:
                    rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[td]")))
                    if i < len(rows):
                        row = rows[i]
                        driver.execute_script("arguments[0].scrollIntoView(true);", row)
                        time.sleep(0.5)
                        row.click()
                        
                        time.sleep(1)
                        
                        try:
                            tools_icon = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "tools")))
                            if tools_icon.is_displayed():
                                print("Found tools icon, processing component...")
                                process_component(driver, wait, download_dir)
                            else:
                                print("Tools icon not visible, skipping component")
                        except Exception as e:
                            print(f"No tools icon found, skipping component: {str(e)}")
                        
                        try:
                            current_url = driver.current_url
                            base_url = current_url.split('#')[0]
                            components_url = base_url + '#/components'
                            driver.get(components_url)
                            time.sleep(2)
                        except Exception as e:
                            print(f"Error navigating to components page: {str(e)}")
                            try:
                                driver.execute_script("window.history.go(-1)")
                                time.sleep(2)
                            except Exception as e2:
                                print(f"Error using history back: {str(e2)}")
                                try:
                                    gradebook_tab = wait.until(EC.element_to_be_clickable((By.ID, "gradebook-tab")))
                                    gradebook_tab.click()
                                    time.sleep(2)
                                    components_button = wait.until(EC.element_to_be_clickable((By.ID, "componentsButton")))
                                    components_button.click()
                                    time.sleep(2)
                                except Exception as e3:
                                    print(f"Failed all navigation attempts: {str(e3)}")
                                    driver.get("https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry#/gradebook")
                                    time.sleep(3)
                                    return
                    else:
                        print(f"Component index {i+1} out of range, skipping")
                
                except Exception as e:
                    print(f"Error finding component rows: {str(e)}")
                    try:
                        current_url = driver.current_url
                        base_url = current_url.split('#')[0]
                        components_url = base_url + '#/components'
                        driver.get(components_url)
                        time.sleep(2)
                    except Exception:
                        return
            
            except Exception as e:
                print(f"Error processing component {i+1}: {str(e)}")
                
                try:
                    driver.get("https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry#/components")
                    time.sleep(2)
                except Exception:
                    driver.get("https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry#/gradebook")
                    time.sleep(2)
                    return
                
    except Exception as e:
        print(f"Error processing components: {str(e)}")
        try:
            screenshot_path = os.path.join(download_dir, "error_screenshot.png")
            driver.save_screenshot(screenshot_path)
            print(f"Error screenshot saved at {screenshot_path}")
        except Exception:
            print("Could not save error screenshot")

def process_component(driver, wait, download_dir):
    try:
        print("  Looking for tools icon...")
        tools_icon = wait.until(EC.element_to_be_clickable((By.ID, "tools")))
        driver.execute_script("arguments[0].scrollIntoView(true);", tools_icon)
        tools_icon.click()
        
        print("  Selecting 'Export Template'...")
        export_template = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='تصدير القالب']")))
        export_template.click()
        
        print("  Selecting 'Excel Spreadsheets'...")
        excel_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='xlsxType' and contains(text(), 'جداول البيانات إكسل')]")))
        excel_option.click()
        
        print("  Clicking 'Export'...")
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@ng-click, 'ok()') and contains(text(), 'تصدير')]")))
        export_button.click()
        
        print("  Waiting for download to complete...")
        wait_for_download(download_dir, 20)
        print("  Download completed successfully.")
    except Exception as e:
        print(f"  Error processing component: {str(e)}")