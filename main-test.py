import os
import time
import logging
import traceback
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

INITIAL_URL = "https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx"
TARGET_URL = "https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/facultyAttendanceTracking#!/markAttendance"

CONTINUE_LINK_CLASS = "export-downlink"
CONTINUE_LINK_SELECTOR = "a.export-downlink"
CONTINUE_LINK_XPATH = "//a[contains(text(), 'متابعة')]"
IMPORT_TABLE_CLASS = "importTable"
VALIDATION_REPORT_SELECTOR = "a[ng-click='generateValidationReport()']"
FINISH_BUTTON_SELECTOR = "a[ng-click='finish()']"
UPLOAD_BUTTON_SELECTOR = "input#upload[value='تحميل']"

def initialize_browser():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_experimental_option("detach", True)
    
    try:
        logger.info("Starting Chrome browser...")
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()
        return driver
        
    except WebDriverException as e:
        logger.error(f"Chrome startup failed: {str(e)}")
        logger.info("Trying Edge browser as fallback...")
        
        try:
            from selenium.webdriver.edge.service import Service as EdgeService
            from selenium.webdriver.edge.options import Options as EdgeOptions
            
            edge_options = EdgeOptions()
            edge_options.add_experimental_option("detach", True)
            driver = webdriver.Edge(options=edge_options)
            driver.maximize_window()
            logger.info("Successfully started Edge browser")
            return driver
            
        except Exception as edge_error:
            logger.error(f"Edge startup also failed: {str(edge_error)}")
            raise

def wait_for_manual_login(driver):
    logger.info("Opening initial website...")
    driver.get(INITIAL_URL)
    initial_handle = driver.current_window_handle
    
    print("\n[MANUAL ACTION REQUIRED]")
    print("Please follow these steps:")
    print("1. Navigate through the website to the login page")
    print("2. Log in with your credentials")
    print("3. Navigate to the markAttendance page which will open in a new tab")
    print("\nThe script will NOT interfere with your browsing.")
    print("Just continue until you reach the markAttendance page in the new tab.")
    
    found_target = False
    max_attempts = 300
    attempts = 0
    
    initial_handles = set(driver.window_handles)
    
    time.sleep(5)
    
    while not found_target and attempts < max_attempts:
        try:
            current_handles = set(driver.window_handles)
            
            new_handles = current_handles - initial_handles
            
            if new_handles:
                for handle in new_handles:
                    try:
                        driver.switch_to.window(handle)
                        
                        tab_url = driver.current_url
                        print(f"Checking new tab: {tab_url}")
                        
                        if TARGET_URL in tab_url:
                            print("\n✅ TARGET PAGE FOUND IN NEW TAB!")
                            print("✅ Staying in this tab and proceeding with import...")
                            found_target = True
                            time.sleep(5)
                            return True
                    except Exception as e:
                        print(f"Error checking tab: {str(e)}")
                        continue
            
            time.sleep(2)
            attempts += 1
            
            if attempts % 15 == 0:
                remaining = max_attempts - attempts
                print(f"Still waiting for target page in a new tab... ({remaining} seconds remaining)")
                new_tab_count = len(current_handles) - len(initial_handles)
                print(f"Currently monitoring {new_tab_count} new tabs")
                
        except Exception as e:
            print(f"Error during window check: {str(e)}")
            time.sleep(2)
            attempts += 1
    
    if not found_target:
        print("Timeout reached before detecting target page")
        return False
        
    return found_target

def navigate_to_import_page(driver):
    logger.info("Navigating to import page...")
    
    try:
        # Step 1: Click on the tools button
        logger.info("Looking for tools button...")
        try:
            tools_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "tools"))
            )
            logger.info("Tools button found!")
        except TimeoutException:
            logger.warning("Tools button not found with ID, trying alternative methods...")
            try:
                tools_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'أدوات') or contains(text(), 'Tools')]"))
                )
                logger.info("Tools button found by text!")
            except TimeoutException:
                tools_buttons = driver.find_elements(By.TAG_NAME, "button")
                for btn in tools_buttons:
                    if btn.get_attribute("id") and ("tool" in btn.get_attribute("id").lower() or "tools" in btn.get_attribute("id").lower()):
                        tools_button = btn
                        logger.info(f"Found potential tools button with id: {btn.get_attribute('id')}")
                        break
                    elif btn.get_attribute("class") and ("tool" in btn.get_attribute("class").lower() or "tools" in btn.get_attribute("class").lower()):
                        tools_button = btn
                        logger.info(f"Found potential tools button with class: {btn.get_attribute('class')}")
                        break
        
        logger.info("Clicking tools button...")
        try:
            # Always use JavaScript for clicking to avoid element not interactable errors
            driver.execute_script("arguments[0].click();", tools_button)
            logger.info("Clicked tools button using JavaScript")
        except Exception as e:
            logger.warning(f"Error clicking tools button: {str(e)}")
            return False
        
        # Wait a moment for the menu to appear
        time.sleep(2)
        
        # Step 2: Click on import tool item
        logger.info("Looking for import tool item or span...")
        
        # First look for the span with "استيراد" text specifically
        import_spans = driver.find_elements(By.XPATH, "//span[contains(text(), 'استيراد')]")
        if import_spans:
            import_item = import_spans[0]
            logger.info("Found import span by text!")
        else:
            # If span not found, try other selectors
            try:
                import_item = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "import-tool-item"))
                )
                logger.info("Import tool item found by ID!")
            except TimeoutException:
                logger.warning("Import item not found with ID, trying alternatives...")
                try:
                    import_item = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, 
                            "//div[contains(@id, 'import-tool-item') or contains(., 'استيراد') or contains(., 'Import')]"))
                    )
                    logger.info("Import item found by xpath!")
                except TimeoutException:
                    menu_items = driver.find_elements(By.CSS_SELECTOR, "li.dropdown-item, li.menu-item, li, div[id*='import']")
                    import_item = None
                    for item in menu_items:
                        try:
                            if "import" in item.text.lower() or "استيراد" in item.text:
                                import_item = item
                                logger.info(f"Found potential import item: {item.text}")
                                break
                        except:
                            pass
                    
                    if not import_item:
                        logger.warning("Could not find import item by typical means, trying direct XPath")
                        # Try more specific XPath expressions
                        import_links = driver.find_elements(By.XPATH, 
                            "//a[contains(text(), 'استيراد')] | //div[contains(text(), 'استيراد')] | //li[contains(text(), 'استيراد')]")
                        if import_links:
                            import_item = import_links[0]
                            logger.info(f"Found import item by direct text match: {import_item.tag_name}")

        if not import_item:
            logger.error("Could not find import item")
            return False

        logger.info("Clicking import tool item...")
        try:
            # Always use JavaScript for clicking to avoid element not interactable errors
            driver.execute_script("arguments[0].click();", import_item)
            logger.info("Clicked import item using JavaScript")
        except Exception as e:
            logger.warning(f"Error clicking import item: {str(e)}")
            return False

        # Wait for import dialog to appear
        logger.info("Waiting for import dialog to appear...")
        time.sleep(3)
        
        return True
    except Exception as e:
        logger.error(f"Error navigating to import page: {str(e)}")
        return False

def wait_for_upload_button_click(driver):
    """
    Wait for the user to upload the Excel file and click the Upload button,
    then attempt to click the continue button when it appears.
    """
    print("\n[MANUAL ACTION REQUIRED]")
    print("Please select your Excel file and click the 'تحميل' (Upload) button.")
    print("The script will detect when the upload is successful and continue automatically...")
    
    max_wait_time = 300  # 5 minutes maximum wait time
    check_interval = 2  # Check every 2 seconds
    elapsed_time = 0
    
    # Store initial state to detect changes
    initial_page_source = driver.page_source
    upload_button_found = False
    
    # First, confirm that we can find the upload button
    try:
        upload_buttons = driver.find_elements(By.CSS_SELECTOR, UPLOAD_BUTTON_SELECTOR)
        if upload_buttons:
            upload_button_found = True
            logger.info(f"Upload button found: {upload_buttons[0].get_attribute('id')}")
    except Exception as e:
        logger.warning(f"Error checking for upload button: {str(e)}")
    
    if not upload_button_found:
        logger.warning("Could not find the upload button. Looking for alternatives...")
        try:
            upload_buttons = driver.find_elements(By.XPATH, "//input[@value='تحميل']")
            if upload_buttons:
                upload_button_found = True
                logger.info(f"Upload button found by value: {upload_buttons[0].get_attribute('id')}")
        except Exception as e:
            logger.warning(f"Error checking for upload button by value: {str(e)}")
    
    if not upload_button_found:
        logger.warning("Upload button not found. Will still monitor for page changes...")
    
    # Main waiting loop
    while elapsed_time < max_wait_time:
        try:
            # Look for continue buttons that may appear after successful upload
            continue_buttons = driver.find_elements(By.CSS_SELECTOR, CONTINUE_LINK_SELECTOR)
            continue_buttons_xpath = driver.find_elements(By.XPATH, CONTINUE_LINK_XPATH)
            
            all_continue_buttons = continue_buttons + [btn for btn in continue_buttons_xpath if btn not in continue_buttons]
            
            # If continue buttons are found, it likely means the upload was successful
            if all_continue_buttons:
                logger.info(f"Found {len(all_continue_buttons)} continue buttons after upload")
                
                # Check if any continue button is clickable
                for button in all_continue_buttons:
                    try:
                        # Store page state before clicking
                        pre_click_url = driver.current_url
                        pre_click_source = driver.page_source
                        
                        # Try to click the continue button
                        logger.info("Attempting to click continue button...")
                        driver.execute_script("arguments[0].click();", button)
                        
                        # Wait briefly to see if anything changes
                        time.sleep(2)
                        
                        # Check if the page changed
                        if driver.current_url != pre_click_url or driver.page_source != pre_click_source:
                            logger.info("Page changed after clicking continue button!")
                            print("\n✅ Upload completed successfully! Moving to the next step...")
                            return True
                    except Exception as click_e:
                        logger.warning(f"Error clicking continue button: {str(click_e)}")
            
            # Check if the page source has significantly changed
            current_page_source = driver.page_source
            if initial_page_source != current_page_source and "uploadSuccess" in current_page_source:
                logger.info("Page content changed and 'uploadSuccess' detected!")
                print("\n✅ Upload completed successfully! Moving to the next step...")
                return True
            
            # Check if the upload button has disappeared (might indicate successful upload)
            try:
                upload_buttons_now = driver.find_elements(By.CSS_SELECTOR, UPLOAD_BUTTON_SELECTOR)
                if upload_button_found and not upload_buttons_now:
                    logger.info("Upload button disappeared after upload")
                    print("\n✅ Upload button disappeared - likely successful! Moving to the next step...")
                    return True
            except Exception as e:
                logger.warning(f"Error checking upload button state: {str(e)}")
            
        except Exception as e:
            logger.warning(f"Error during upload detection: {str(e)}")
        
        time.sleep(check_interval)
        elapsed_time += check_interval
        
        # Show progress message every 30 seconds
        if elapsed_time % 30 == 0:
            remaining = max_wait_time - elapsed_time
            print(f"Still waiting for file upload... ({remaining} seconds remaining)")
    
    print("\n❌ Timeout reached. Could not detect file upload.")
    return False

def click_continue_buttons_sequence(driver):
    """Attempt to click all continue buttons in sequence"""
    try:
        # Try to click any continue buttons that are available
        for i in range(5):  # Try up to 5 continue buttons
            logger.info(f"Looking for continue button (attempt {i+1}/5)")
            
            # Find continue buttons
            continue_buttons = driver.find_elements(By.CSS_SELECTOR, CONTINUE_LINK_SELECTOR)
            xpath_continue_buttons = driver.find_elements(By.XPATH, CONTINUE_LINK_XPATH)
            
            all_continue_buttons = continue_buttons + [btn for btn in xpath_continue_buttons if btn not in continue_buttons]
            
            if not all_continue_buttons:
                logger.info(f"No continue buttons found on attempt {i+1}")
                break
            
            # Current state before clicking
            current_url = driver.current_url
            current_page_state = driver.page_source
            
            # Click the button using JavaScript
            try:
                logger.info(f"Clicking continue button (attempt {i+1})")
                driver.execute_script("arguments[0].click();", all_continue_buttons[0])
                
                # Wait for page change
                time.sleep(3)
                
                # Check if the page changed
                if driver.current_url != current_url or driver.page_source != current_page_state:
                    logger.info(f"Page changed after clicking continue button (attempt {i+1})")
                else:
                    logger.info(f"Page did not change after clicking continue button (attempt {i+1})")
                    # If the page didn't change, it might be the end of the sequence
                    # or the button might not be active yet
                    time.sleep(2)  # Wait a bit longer and see if anything changes
                    if driver.page_source == current_page_state:
                        logger.info("Page still hasn't changed, attempting to find next active button")
            except Exception as e:
                logger.warning(f"Error clicking continue button (attempt {i+1}): {str(e)}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error in click_continue_buttons_sequence: {str(e)}")
        return False

def complete_import_process(driver):
    try:
        # First click any remaining continue buttons
        click_continue_buttons_sequence(driver)
        
        # Look for the import table
        logger.info("Checking for import table...")
        try:
            import_table = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, IMPORT_TABLE_CLASS))
            )
            logger.info("Import table found!")
        except TimeoutException:
            logger.warning("Import table not found with class name, trying alternatives...")
            try:
                import_tables = driver.find_elements(By.CSS_SELECTOR, "table, div.table")
                if import_tables:
                    import_table = import_tables[0]
                    logger.info(f"Found potential import table: {import_table.get_attribute('class')}")
                else:
                    logger.warning("Could not find any import table, continuing anyway...")
            except Exception as e:
                logger.warning(f"Error finding import table alternative: {str(e)}")
        
        # Look for the validation report link
        logger.info("Looking for validation report link...")
        try:
            validation_report_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, VALIDATION_REPORT_SELECTOR))
            )
            logger.info("Validation report link found!")
        except TimeoutException:
            logger.warning("Validation report link not found with selector, trying alternatives...")
            try:
                validation_report_links = driver.find_elements(By.XPATH, 
                    "//a[contains(text(), 'تنزيل تقرير التحقق') or contains(text(), 'Download validation report')]")
                if validation_report_links:
                    validation_report_link = validation_report_links[0]
                    logger.info(f"Found validation report link by text: {validation_report_link.text}")
                else:
                    logger.warning("Could not find any validation report link, continuing anyway...")
                    validation_report_link = None
            except Exception as e:
                logger.warning(f"Error finding validation report link alternative: {str(e)}")
                validation_report_link = None
        
        # Click the validation report link if found
        if validation_report_link:
            logger.info("Clicking validation report link...")
            try:
                driver.execute_script("arguments[0].click();", validation_report_link)
                logger.info("Validation report link clicked successfully")
            except Exception as e:
                logger.warning(f"Error clicking validation report link: {str(e)}")
            
            # Wait for report to download
            logger.info("Waiting for validation report to download (5 seconds)...")
            time.sleep(5)
        
        # Look for the finish button
        logger.info("Looking for finish button...")
        try:
            finish_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, FINISH_BUTTON_SELECTOR))
            )
            logger.info("Finish button found!")
        except TimeoutException:
            logger.warning("Finish button not found with selector, trying alternatives...")
            try:
                finish_buttons = driver.find_elements(By.XPATH, 
                    "//a[contains(text(), 'إنهاء') or contains(text(), 'Finish')]")
                if finish_buttons:
                    finish_button = finish_buttons[0]
                    logger.info(f"Found finish button by text: {finish_button.text}")
                else:
                    logger.warning("Could not find any finish button")
                    return False
            except Exception as e:
                logger.warning(f"Error finding finish button alternative: {str(e)}")
                return False
        
        # Click the finish button
        logger.info("Clicking finish button...")
        try:
            driver.execute_script("arguments[0].click();", finish_button)
            logger.info("Finish button clicked successfully")
        except Exception as e:
            logger.warning(f"Error clicking finish button: {str(e)}")
        
        # Wait for completion
        logger.info("Waiting for completion (3 seconds)...")
        time.sleep(3)
        
        logger.info("Excel file import process completed successfully!")
        return True
        
    except Exception as e:
        logger.error(f"Error during import process completion: {str(e)}")
        return False

def main():
    print("\n=== Starting Excel File Import System ===\n")
    
    driver = None
    try:
        driver = initialize_browser()

        print("\nStep 1: Waiting for you to reach the markAttendance page in a new tab...")
        login_success = wait_for_manual_login(driver)
        
        if login_success:
            print("\n▶▶▶ Step 1 COMPLETE: markAttendance page found! ◀◀◀")
            print("Waiting 5 seconds for page to fully load...")
            time.sleep(5)

            print("\n▶▶▶ Step 2: Navigating to import page... ◀◀◀")
            navigation_success = navigate_to_import_page(driver)
            
            if navigation_success:
                print("\n▶▶▶ Step 2 COMPLETE: Import page opened! ◀◀◀")
                print("\n▶▶▶ Step 3: Waiting for manual file upload... ◀◀◀")
                
                upload_success = wait_for_upload_button_click(driver)
                
                if upload_success:
                    print("\n▶▶▶ Step 3 COMPLETE: File uploaded successfully! ◀◀◀")
                    print("\n▶▶▶ Step 4: Completing import process... ◀◀◀")
                    
                    import_success = complete_import_process(driver)
                    
                    if import_success:
                        print("\n▶▶▶ Step 4 COMPLETE: Import process finished successfully! ◀◀◀")
                        print("✅ Excel file has been imported. Check the validation report for details.")
                    else:
                        print("\n❌ Import process completion failed.")
                        print("You may need to complete the remaining steps manually.")
                else:
                    print("\n❌ No file upload detected within the timeout period.")
                    print("Please try again or complete the process manually.")
            else:
                print("\n❌ Navigation to import page failed.")
                print("Please try to navigate to the import page manually.")
        else:
            print("\n❌ Login timeout or error occurred.")
            print("Please try again.")
    
    except Exception as e:
        print("\n❌ An unexpected error occurred:")
        print(str(e))
        traceback.print_exc()
    
    finally:
        if driver:
            print("\nNote: Browser has been left open for your use.")
            print("You can close it manually when finished.")
        
        print("\n=== Excel File Import System Finished ===")

if __name__ == "__main__":
    main()