import os
import time
import logging
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException

from .config import (
    TOOLS_BUTTON_ID,
    EXPORT_TOOL_ITEM_ID,
    OUTPUT_FILENAME
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def export_attendance_data(driver):
    try:
        # Step 1: Click on the tools button
        logger.info("Looking for tools button...")
        try:
            tools_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, TOOLS_BUTTON_ID))
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
                    if "tool" in btn.get_attribute("id").lower() or "tools" in btn.get_attribute("class").lower():
                        tools_button = btn
                        logger.info(f"Found potential tools button with id: {btn.get_attribute('id')}")
                        break
        
        logger.info("Clicking tools button...")
        try:
            tools_button.click()
            logger.info("Tools button clicked successfully")
        except Exception as e:
            logger.warning(f"Error clicking tools button directly: {str(e)}")
            driver.execute_script("arguments[0].click();", tools_button)
            logger.info("Clicked tools button using JavaScript")
        
        # Step 2: Click on export template item
        logger.info("Looking for export template item...")
        try:
            export_item = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, EXPORT_TOOL_ITEM_ID))
            )
            logger.info("Export template item found!")
        except TimeoutException:
            logger.warning("Export item not found with ID, trying alternatives...")
            try:
                export_item = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        "//li[contains(@class, 'dropdown-item') and (contains(text(), 'تصدير') or contains(text(), 'Export'))]"))
                )
                logger.info("Export item found by text!")
            except TimeoutException:
                menu_items = driver.find_elements(By.CSS_SELECTOR, "li.dropdown-item, li.menu-item, li")
                for item in menu_items:
                    try:
                        if "export" in item.text.lower() or "تصدير" in item.text:
                            export_item = item
                            logger.info(f"Found potential export item: {item.text}")
                            break
                    except:
                        pass

        logger.info("Clicking export template item...")
        try:
            export_item.click()
            logger.info("Export item clicked successfully")
        except Exception as e:
            logger.warning(f"Error clicking export item directly: {str(e)}")
            driver.execute_script("arguments[0].click();", export_item)
            logger.info("Clicked export item using JavaScript")

        # Wait for export dialog to appear
        logger.info("Waiting for export dialog to appear...")
        time.sleep(5)  # Increased wait time to make sure dialog fully loads
        
        # ---- IMPROVED EXCEL FORMAT SELECTION ----
        logger.info("Selecting Excel format using EXACT selectors...")
        
        # Try multiple methods to find and click the Excel format option
        excel_format_found = False
        
        # Method 1: Using the exact class and for attribute
        try:
            logger.info("Method 1: Using exact CSS selector for Excel format...")
            excel_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 
                    "label.export-innerLabel.export-commonLabel[for='xlsxType']"))
            )
            logger.info("Excel format found using exact CSS selector!")
            
            # Use multiple click methods
            try:
                excel_label.click()
                logger.info("Excel format clicked normally")
                excel_format_found = True
            except:
                driver.execute_script("arguments[0].click();", excel_label)
                logger.info("Excel format clicked using JavaScript")
                excel_format_found = True
        except Exception as e:
            logger.warning(f"Method 1 failed: {str(e)}")
        
        # Method 2: Using exact text content
        if not excel_format_found:
            try:
                logger.info("Method 2: Using exact text content for Excel format...")
                excel_label = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                        "//label[contains(text(), 'جداول البيانات إكسل (.Xlsx)')]"))
                )
                logger.info("Excel format found using exact text content!")
                
                driver.execute_script("arguments[0].click();", excel_label)
                logger.info("Excel format clicked using JavaScript")
                excel_format_found = True
            except Exception as e:
                logger.warning(f"Method 2 failed: {str(e)}")
        
        # Method 3: Find label by ID then find its associated radio button and click that
        if not excel_format_found:
            try:
                logger.info("Method 3: Using ID to find and click Excel radio button...")
                excel_radio = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "xlsxType"))
                )
                logger.info("Excel radio button found by ID!")
                
                driver.execute_script("arguments[0].checked = true; arguments[0].click();", excel_radio)
                logger.info("Excel radio button set to checked state using JavaScript")
                excel_format_found = True
            except Exception as e:
                logger.warning(f"Method 3 failed: {str(e)}")
        
        # Method 4: Try to find any radio button or label that mentions Excel
        if not excel_format_found:
            try:
                logger.info("Method 4: Looking for any element mentioning Excel...")
                excel_elements = driver.find_elements(By.XPATH, 
                    "//*[contains(text(), 'Excel') or contains(text(), 'إكسل') or contains(text(), 'xlsx')]")
                
                if excel_elements:
                    logger.info(f"Found {len(excel_elements)} potential Excel elements")
                    for element in excel_elements:
                        try:
                            element_type = element.tag_name
                            element_text = element.text
                            logger.info(f"Trying to click Excel element: {element_type} with text '{element_text}'")
                            driver.execute_script("arguments[0].click();", element)
                            excel_format_found = True
                            logger.info("Excel element clicked successfully")
                            break
                        except:
                            continue
            except Exception as e:
                logger.warning(f"Method 4 failed: {str(e)}")
        
        # Log the elements and print the page source for debugging if Excel format still not found
        if not excel_format_found:
            logger.warning("All methods failed to find Excel format option!")
            logger.info("Dumping all labels on the page for debugging:")
            try:
                labels = driver.find_elements(By.TAG_NAME, "label")
                for idx, label in enumerate(labels):
                    try:
                        label_for = label.get_attribute("for")
                        label_text = label.text
                        logger.info(f"Label {idx}: for='{label_for}', text='{label_text}'")
                    except:
                        pass
            except:
                pass
        else:
            logger.info("Excel format option selected successfully!")
        
        # ---- END OF IMPROVED EXCEL FORMAT SELECTION ----
        
        # Step 4: Select "All students in section" option
        logger.info("Selecting all students option...")
        try:
            all_students_label = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, 
                    "//label[@for='allStudRadio' or contains(text(), 'قائمة بكافة طلاب الشعبة')]"))
            )
            logger.info("All students option found!")
            driver.execute_script("arguments[0].click();", all_students_label)
            logger.info("All students option selected")
        except Exception as e:
            logger.warning(f"Error selecting all students option: {str(e)}")
            try:
                all_students_radio = driver.find_element(By.ID, "allStudRadio")
                driver.execute_script("arguments[0].click();", all_students_radio)
                logger.info("All students option selected using radio button directly")
            except:
                logger.warning("Could not find all students option. Continuing with default.")
        
        # Step 5: Select "All lecture days" option
        logger.info("Selecting all lecture days option...")
        try:
            all_days_label = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, 
                    "//label[@for='dateRange' or contains(text(), 'كل أيام المحاضرات')]"))
            )
            logger.info("All lecture days option found!")
            driver.execute_script("arguments[0].click();", all_days_label)
            logger.info("All lecture days option selected")
        except Exception as e:
            logger.warning(f"Error selecting all lecture days option: {str(e)}")
            try:
                all_days_radio = driver.find_element(By.ID, "dateRange")
                driver.execute_script("arguments[0].click();", all_days_radio)
                logger.info("All lecture days option selected using radio button directly")
            except:
                logger.warning("Could not find all lecture days option. Continuing with default.")
        
        # Step 6: Click export button
        logger.info("Looking for export button...")
        export_button = None
        try:
            # First try the link with "تصدير" text
            export_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'تصدير')]"))
            )
            logger.info("Export button found (link with تصدير text)!")
        except TimeoutException:
            logger.warning("Export button not found as link with تصدير text, trying alternatives...")
            
            # Try by CSS selector for export-downlink
            try:
                export_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a.export-downlink"))
                )
                logger.info("Export button found (a.export-downlink)!")
            except TimeoutException:
                logger.warning("Export button not found with a.export-downlink, trying more alternatives...")
                
                # Try generic button or link with export/تصدير text
                try:
                    export_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'تصدير') or contains(text(), 'Export')]"))
                    )
                    logger.info("Export button found (generic element with تصدير text)!")
                except TimeoutException:
                    # Try finding ng-click attribute
                    try:
                        export_button = driver.find_element(By.CSS_SELECTOR, "[ng-click='ok()']")
                        logger.info("Export button found (element with ng-click='ok()')!")
                    except:
                        logger.error("No export button found with any method!")
        
        if export_button:
            logger.info("Clicking export button...")
            try:
                export_button.click()
                logger.info("Export button clicked successfully")
            except Exception as e:
                logger.warning(f"Error clicking export button directly: {str(e)}")
                try:
                    driver.execute_script("arguments[0].click();", export_button)
                    logger.info("Clicked export button using JavaScript")
                except Exception as js_e:
                    logger.error(f"Failed to click export button even with JavaScript: {str(js_e)}")
                    
                    # Last resort: try to find the OK button by text if available
                    try:
                        ok_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Ok') or contains(text(), 'ok') or contains(text(), 'تأكيد') or contains(text(), 'موافق')]")
                        if ok_buttons:
                            driver.execute_script("arguments[0].click();", ok_buttons[0])
                            logger.info("Clicked OK button as last resort")
                    except:
                        return False
        else:
            logger.error("No export button found, cannot continue")
            return False
        
        # Wait for file download
        logger.info("Waiting for file download (15 seconds)...")
        time.sleep(15)
        
        # Process downloaded file
        logger.info(f"Looking for downloaded file to rename to {OUTPUT_FILENAME}...")
        download_dir = str(Path.home() / "Downloads")
        desktop_dir = str(Path.home() / "Desktop")
        
        try:
            # Log all files in downloads directory for debugging
            download_files = os.listdir(download_dir)
            logger.info(f"Files in downloads folder ({len(download_files)} total):")
            
            # Get files created/modified in the last minute
            recent_files = []
            for file_path in [os.path.join(download_dir, f) for f in download_files]:
                file_time = os.path.getmtime(file_path)
                if time.time() - file_time < 60:  # If modified in the last minute
                    recent_files.append(file_path)
            
            logger.info(f"Recent files (modified in last minute): {len(recent_files)}")
            for idx, file_path in enumerate(recent_files, 1):
                file_name = os.path.basename(file_path)
                file_time = time.ctime(os.path.getmtime(file_path))
                logger.info(f"{idx}. {file_name} (Modified: {file_time})")
            
            # Look for excel files first
            excel_files = [f for f in recent_files if f.endswith('.xlsx') or f.endswith('.xls')]
            if excel_files:
                # Use the most recent Excel file
                latest_file = max(excel_files, key=os.path.getmtime)
            elif recent_files:
                # If no Excel files, use the most recent file of any type
                latest_file = max(recent_files, key=os.path.getmtime)
            else:
                # If no recent files, use the most recent file in the downloads folder
                latest_file = max([os.path.join(download_dir, f) for f in download_files], key=os.path.getmtime)
            
            file_name = os.path.basename(latest_file)
            logger.info(f"Found file to use: {file_name}")
            
            # Create target path
            target_path = os.path.join(desktop_dir, OUTPUT_FILENAME)
            
            # Check if target file already exists and remove it
            if os.path.exists(target_path):
                logger.info(f"Target file already exists, removing it: {target_path}")
                os.remove(target_path)
            
            # Copy file to desktop
            logger.info(f"Moving file from {latest_file} to {target_path}")
            os.rename(latest_file, target_path)
            logger.info(f"File successfully saved to: {target_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error processing downloaded file: {str(e)}")
            return False
            
    except Exception as e:
        logger.error(f"Error during export process: {str(e)}")
        return False