MyTVTC - TVTC Attendance and Grades Management System
A comprehensive solution for managing student attendance and grades at the Technical and Vocational Training Corporation (TVTC). This integrated system automates attendance tracking using QR codes, manages grade templates, and provides an efficient interface for educators.
Features
1. QR-based Attendance System

Generate QR codes for students to scan for attendance
Web-based scanning interface accessible from any device
Secure attendance validation with device fingerprinting
Automatic absence marking after session expiration
Export attendance data to Excel format

2. Grade Management

Extract grade templates from TVTC portal
Process and merge Excel files with grade data
Split files by component (LEC1, LEC2, LAB1, LAB2)
Update grade records automatically

3. Data Export and Automation

Automate browser interactions with TVTC website
Export attendance data to desktop
Simplified login and navigation process

4. Unified User Interface

Modern, intuitive interface for all system functions
Real-time tracking of attendance sessions
Centralized access to all tools and features

System Requirements

Windows 10/11
Python 3.8 or higher
Chrome browser
Active internet connection
TVTC faculty account

Installation

Clone or download this repository to your local machine
Install required dependencies:

bashpip install -r requirements.txt

Run the main application:

bashpython main_ui.py
Module Structure

main_ui.py: Main application interface
main-qr_attendance.py: QR code attendance system
main-Grad.py: Grades processing system
main-scraping.py: Data export system
main-finished.py: Simplified browser opener

Key Directories

project/: Contains browser automation, Excel processing, and file validation
qr_attendance/: QR code generation, web server, and Excel handling
scraping/: Web scraping and data fetching
security/: Device fingerprinting for attendance verification

Usage Instructions
Attendance Tracking

Click "بدء التحضير" (Start Attendance) in the main interface
Select the Excel file containing student information
Set the attendance duration
Show the generated QR code to students
Students scan the QR code with their mobile devices and enter their information
The system automatically records attendance and marks absent students after the session expires

Grade Management

Click "تصدير قالب التحضير" (Export Attendance Template) to extract data from TVTC
Log in to your TVTC account when prompted
The system will automatically navigate to the required pages and export data
Click "نظام رصد الدرجات" (Grade Monitoring System) to process grades
Select the grade template Excel file
The system will process the file and split it by component
Updated files will be saved to your desktop

Uploading Attendance

Click "رفع التحضير في الموقع" (Upload Attendance to Website)
The system will open a browser window to the TVTC portal
Complete the upload process through the TVTC interface

Security Features

Attendance verification using device fingerprinting
Prevention of attendance fraud with unique device identification
Secure storage of device fingerprints
Validation checks for student identity

Troubleshooting

If browser automation fails, try manually logging in through the opened browser
Ensure Excel files are in the correct format with required columns
Check internet connection if QR code scanning fails
For device fingerprint issues, students should use their own devices

Development
The system is built using:

Python for core functionality
Selenium for browser automation
Flask for the web server
Pandas for Excel processing
TKinter for the user interface
QR code generation and scanning libraries

License
Developed by Eng. Saad Yasser Al-Ali
For technical support: 0557000646
Version
Current version: 2.1.0