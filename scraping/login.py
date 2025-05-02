import time
import os
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchWindowException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from .config import INITIAL_URL, TARGET_URL

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
    
    # Initial window handles at the start
    initial_handles = set(driver.window_handles)
    
    # Wait for initial page to fully load before starting to check
    time.sleep(5)
    
    while not found_target and attempts < max_attempts:
        try:
            # Get current window handles
            current_handles = set(driver.window_handles)
            
            # Only look for target page in NEW tabs (opened after script started)
            new_handles = current_handles - initial_handles
            
            # Check all new tabs first
            if new_handles:
                for handle in new_handles:
                    try:
                        # Switch to this new tab
                        driver.switch_to.window(handle)
                        
                        # Get the URL of this tab
                        tab_url = driver.current_url
                        print(f"Checking new tab: {tab_url}")
                        
                        if TARGET_URL in tab_url:
                            print("\n✅ TARGET PAGE FOUND IN NEW TAB!")
                            print("✅ Staying in this tab and proceeding with export...")
                            found_target = True
                            # Explicitly wait here to ensure page is loaded
                            time.sleep(5)
                            return True
                    except Exception as e:
                        print(f"Error checking tab: {str(e)}")
                        continue
            
            # Don't go back to the initial tab - stay wherever we last checked
            
            # Sleep between checks
            time.sleep(2)
            attempts += 1
            
            # Show progress message every 15 seconds
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