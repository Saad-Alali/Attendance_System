import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

TVTC_URL = "https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx"

def main():
    print("\n=== Simplified TVTC Browser Opener ===\n")
    
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("detach", True)
    
    print("Starting Chrome browser...")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    
    print(f"Opening URL: {TVTC_URL}")
    driver.get(TVTC_URL)
    
    print("Browser opened. Program will wait until browser is closed by user...")
    
    try:
        while driver.window_handles:
            time.sleep(1)
    except:
        pass
    
    print("\n=== Browser closed. Script finished ===")

if __name__ == "__main__":
    main()