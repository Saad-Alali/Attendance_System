import pyautogui
import time
import os

pyautogui.FAILSAFE = True  # Move mouse to top-left corner to abort

def click_at_position(x, y, description="specific point"):
    """Click at a specific position on screen with description for clarity"""
    print(f"Clicking at {description} - Position ({x}, {y})")
    pyautogui.moveTo(x, y, duration=0.5)  # Move mouse to position with animation
    pyautogui.click()
    time.sleep(1)  # Wait after clicking

def find_coordinates_helper():
    """Helper function to find screen coordinates"""
    print("Coordinate finder helper started...")
    print("Move your mouse to desired positions and note the coordinates.")
    print("Press Ctrl+C to exit this helper function.")
    
    try:
        while True:
            x, y = pyautogui.position()
            position_str = f"Current mouse position: X: {x} Y: {y}"
            print(position_str, end='\r')
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nCoordinate finder exited.")

def upload_file(file_path, upload_button_x, upload_button_y):
    """Upload a file by clicking on an upload button and entering the file path"""
    # Ensure file exists
    if not os.path.exists(file_path):
        print(f"Error: File {file_path} does not exist!")
        return False
    
    # Click the upload button
    click_at_position(upload_button_x, upload_button_y, "upload button")
    time.sleep(2)  # Wait for file dialog to open
    
    # Type the file path
    pyautogui.write(file_path)
    time.sleep(1)
    
    # Press Enter to confirm
    pyautogui.press('enter')
    time.sleep(3)  # Wait for file to upload
    
    print(f"File {file_path} uploaded successfully.")
    return True

def navigate_website_and_upload_example():
    """Example function to navigate a website and upload a file"""
    # Define the coordinates for your specific website
    # These are example coordinates - you should replace them with actual values
    # Use the find_coordinates_helper() function to determine these values
    
    # Website navigation points
    BROWSER_ADDRESS_X, BROWSER_ADDRESS_Y = 400, 50  # Browser address bar
    LOGIN_BUTTON_X, LOGIN_BUTTON_Y = 800, 500       # Login button
    USERNAME_FIELD_X, USERNAME_FIELD_Y = 500, 300   # Username field
    PASSWORD_FIELD_X, PASSWORD_FIELD_Y = 500, 350   # Password field
    UPLOAD_PAGE_BUTTON_X, UPLOAD_PAGE_BUTTON_Y = 200, 200  # Button to navigate to upload page
    UPLOAD_BUTTON_X, UPLOAD_BUTTON_Y = 500, 400     # Upload button
    SUBMIT_BUTTON_X, SUBMIT_BUTTON_Y = 500, 600     # Submit button after upload
    
    # File to upload (replace with your actual file path)
    FILE_PATH = "C:\\path\\to\\your\\file.pdf"  # Use double backslashes or raw string r"C:\path\to\file.pdf"
    
    # Open the website (assuming browser is already open)
    click_at_position(BROWSER_ADDRESS_X, BROWSER_ADDRESS_Y, "browser address bar")
    pyautogui.hotkey('ctrl', 'a')  # Select all text in the address bar
    time.sleep(0.5)
    pyautogui.write("https://example.com")  # Replace with your website
    pyautogui.press('enter')
    
    # Wait for website to load
    print("Waiting for website to load...")
    time.sleep(5)
    
    # Login (if needed)
    click_at_position(USERNAME_FIELD_X, USERNAME_FIELD_Y, "username field")
    pyautogui.write("your_username")  # Replace with your username
    
    click_at_position(PASSWORD_FIELD_X, PASSWORD_FIELD_Y, "password field")
    pyautogui.write("your_password")  # Replace with your password
    
    click_at_position(LOGIN_BUTTON_X, LOGIN_BUTTON_Y, "login button")
    
    # Wait for login to complete
    print("Waiting for login to complete...")
    time.sleep(5)
    
    # Navigate to upload page
    click_at_position(UPLOAD_PAGE_BUTTON_X, UPLOAD_PAGE_BUTTON_Y, "upload page button")
    time.sleep(3)
    
    # Upload the file
    upload_file(FILE_PATH, UPLOAD_BUTTON_X, UPLOAD_BUTTON_Y)
    
    # Submit the form
    click_at_position(SUBMIT_BUTTON_X, SUBMIT_BUTTON_Y, "submit button")
    
    print("Process completed successfully!")

def main_menu():
    while True:
        print("\n" + "=" * 50)
        print("PyAutoGUI Precise Click and Upload Tool")
        print("=" * 50)
        print("1. Run coordinate finder helper")
        print("2. Click at specific position")
        print("3. Upload a file to specific position")
        print("4. Run website navigation and upload example")
        print("0. Exit")
        
        choice = input("\nSelect an option: ")
        
        if choice == '1':
            find_coordinates_helper()
        elif choice == '2':
            try:
                x = int(input("Enter X coordinate: "))
                y = int(input("Enter Y coordinate: "))
                description = input("Enter description (optional): ") or "specified position"
                click_at_position(x, y, description)
            except ValueError:
                print("Invalid coordinates. Please enter numbers only.")
        elif choice == '3':
            try:
                file_path = input("Enter full file path: ")
                x = int(input("Enter upload button X coordinate: "))
                y = int(input("Enter upload button Y coordinate: "))
                upload_file(file_path, x, y)
            except ValueError:
                print("Invalid coordinates. Please enter numbers only.")
        elif choice == '4':
            print("Warning: This will attempt to navigate a website and perform actions.")
            print("You need to modify the coordinates in the code to match your specific website.")
            confirm = input("Do you want to continue? (y/n): ")
            if confirm.lower() == 'y':
                navigate_website_and_upload_example()
        elif choice == '0':
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    print("Welcome to PyAutoGUI Precise Click and Upload Tool!")
    print("WARNING: Move mouse to top-left corner anytime to emergency stop the program.")
    
    main_menu()