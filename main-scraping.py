import time
import logging
import traceback
from scraping.login import initialize_browser, wait_for_manual_login
from scraping.fetch_data import export_attendance_data

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    print("\n=== Starting Attendance Data Export System ===\n")
    
    driver = None
    try:
        driver = initialize_browser()

        print("\nStep 1: Waiting for you to reach the markAttendance page in a new tab...")
        login_success = wait_for_manual_login(driver)
        
        if login_success:
            print("\n▶▶▶ Step 1 COMPLETE: markAttendance page found! ◀◀◀")
            print("▶▶▶ MOVING TO STEP 2 NOW... ◀◀◀")

            print("Waiting 5 seconds for page to fully load...")
            time.sleep(5)

            print("\n▶▶▶ Step 2 STARTING: Exporting attendance data... ◀◀◀")
            print("Looking for tools button to begin export process...")

            export_success = export_attendance_data(driver)
            
            if export_success:
                print("\n▶▶▶ Step 2 COMPLETE: Export Successful! ◀◀◀")
                print("✅ Attendance data has been exported to your desktop as SAAD.xlsx")
            else:
                print("\n❌ Export process failed.")
                print("You can try again or export manually.")
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
        
        print("\n=== Attendance Data Export System Finished ===")

if __name__ == "__main__":
    main()