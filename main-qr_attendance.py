import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import threading
import time
import sys
import io
import pandas as pd
import webbrowser
from pyngrok import ngrok
from pyngrok import conf as ngrok_conf

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from qr_attendance.excel_handler import ExcelHandler
from qr_attendance.generate_qr import generate_lecture_qr
from qr_attendance.web_server import start_server
from qr_attendance.config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    
    return file_path if file_path else None

def attendance_timer(excel_handler, lecture_date, attendance_duration=ATTENDANCE_DURATION):
    print(f"\nStarting attendance session for {attendance_duration//60} minute(s)...")
    time.sleep(attendance_duration)
    
    print("\n=== Attendance time expired ===")
    absent_count = excel_handler.mark_all_absent(lecture_date)
    print(f"{absent_count} students marked as absent")
    
    success, message = excel_handler.save_file()
    print(message)
    
    print("\nAttendance session closed and data saved")
    
    # حذف ملفات الإخراج بعد انتهاء مدة التحضير
    try:
        output_files = [os.path.abspath(OUTPUT_FILENAME), os.path.abspath(QR_OUTPUT_FILENAME)]
        for file_path in output_files:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"تم حذف الملف: {file_path}")
    except Exception as e:
        print(f"خطأ أثناء حذف الملفات: {e}")
    
    print("\nProgram will close in 3 seconds...")
    time.sleep(3)
    os._exit(0)

def mark_student_attendance(excel_handler, lecture_date, student_name, student_id, validate_only=False):
    full_name_col = COLUMN_NAMES["full_name"]
    student_id_col = COLUMN_NAMES["student_id"]
    date_col = COLUMN_NAMES["date"]
    
    # Check if the student is registered for today's lecture
    student_mask = (
        (excel_handler.df[student_id_col].astype(str).str.strip() == str(student_id).strip()) &
        (excel_handler.df[date_col].dt.date == lecture_date.date())
    )
    
    if not any(student_mask):
        return False, "خطأ: الطالب غير مسجل في محاضرة اليوم"
    
    if validate_only:
        # Just return validation result, don't mark attendance
        return True, "الطالب موجود في قائمة المحاضرة"
    
    # Proceed with marking attendance
    success, message = excel_handler.mark_attendance(student_name, student_id, lecture_date)
    print(f"Web submission: {student_name} ({student_id}) - {message}")
    return success, message

def reset_previous_attendance_data(excel_handler, lecture_date):
    count = excel_handler.reset_attendance_for_date(lecture_date)
    print(f"Reset attendance data for {count} students for date {lecture_date.strftime('%Y-%m-%d')}")
    
    # Save the file after reset
    success, message = excel_handler.save_file()
    if success:
        print("File saved successfully after resetting attendance data")
    else:
        print(f"Warning: {message}")
    
    return count

def main():
    try:
        print("=== QR Attendance System (Web Version) ===")
        
        # Configure ngrok
        ngrok_conf.get_default().authtoken = '2uJbHU37cEkmp2othhPYSZjNY8N_5trhFGNwJKjswwSp89k6Y'
        
        excel_file = select_excel_file()
        if not excel_file:
            print("No file selected. Exiting program.")
            return
        
        excel_handler = ExcelHandler(excel_file)
        excel_handler.load_file()
        
        # Debug info - show a sample of the date column before conversion
        date_col = COLUMN_NAMES["date"]
        if date_col in excel_handler.df.columns:
            sample_dates = excel_handler.df[date_col].head(3).tolist()
            print(f"Sample dates before conversion: {sample_dates}")
        
        # Convert date column with improved handling
        excel_handler.convert_date_column()
        
        # Debug info - show a sample after conversion
        if date_col in excel_handler.df.columns:
            try:
                sample_dates = excel_handler.df[date_col].head(3).dt.strftime('%Y-%m-%d').tolist()
                print(f"Sample dates after conversion: {sample_dates}")
            except:
                print("Could not format converted dates")
        
        today_lectures = excel_handler.check_lecture_today()
        if today_lectures is None or today_lectures.empty:
            print("No lecture scheduled for today.")
            # Show what today's date is for debugging
            today = pd.Timestamp.now().date()
            print(f"Today's date is: {today}")
            
            # عرض رسالة للمستخدم إذا لم تكن هناك محاضرة اليوم
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("لا توجد محاضرة", "لا توجد محاضرة مجدولة لليوم الحالي.")
            
            # Show all unique dates in the file for debugging
            if date_col in excel_handler.df.columns:
                try:
                    unique_dates = sorted(excel_handler.df[date_col].dt.date.unique())
                    print(f"Available lecture dates in file: {unique_dates}")
                except:
                    print("Could not extract unique dates")
            return
        
        lecture_info = today_lectures.iloc[0]
        
        lecture_name_col = COLUMN_NAMES["lecture_name"]
        section_col = COLUMN_NAMES["section"]
        date_col = COLUMN_NAMES["date"]
        
        lecture_name = "Unknown Lecture"
        if lecture_name_col in lecture_info.index and not pd.isnull(lecture_info[lecture_name_col]):
            lecture_name = lecture_info[lecture_name_col]
        elif section_col in lecture_info.index and not pd.isnull(lecture_info[section_col]):
            lecture_name = lecture_info[section_col]
            
        lecture_date = lecture_info[date_col]
        
        print(f"\nLecture found for today: {lecture_name} on {lecture_date.strftime('%Y-%m-%d')}")
        
        # Reset previous attendance data for today's lecture
        print("\n=== Resetting previous attendance data ===")
        reset_previous_attendance_data(excel_handler, lecture_date)
        
        # طلب مدة التحضير من المستخدم
        root = tk.Tk()
        root.withdraw()
        
        # القيمة الافتراضية هي 15 دقيقة
        default_duration = ATTENDANCE_DURATION // 60
        
        # طلب مدة التحضير من المستخدم
        duration_minutes = simpledialog.askinteger(
            "مدة التحضير", 
            "أدخل مدة التحضير بالدقائق:",
            initialvalue=default_duration,
            minvalue=1,
            maxvalue=120
        )
        
        # إذا ألغى المستخدم، استخدم القيمة الافتراضية
        if duration_minutes is None:
            duration_minutes = default_duration
            print(f"تم استخدام مدة التحضير الافتراضية: {duration_minutes} دقيقة")
        else:
            print(f"تم تحديد مدة التحضير: {duration_minutes} دقيقة")
        
        # تحويل الدقائق إلى ثواني
        custom_attendance_duration = duration_minutes * 60
        
        session_code = str(int(time.time()))[-8:]
        
        # Open ngrok tunnel
        try:
            # Create a tunnel with simplified parameters
            ngrok_tunnel = ngrok.connect(5000)  # Just use the port number directly
            public_url = ngrok_tunnel.public_url
            
            print(f"\nNgrok Public URL: {public_url}")
        except Exception as ngrok_error:
            print(f"Error creating ngrok tunnel: {ngrok_error}")
            messagebox.showerror("Ngrok Error", f"Failed to create ngrok tunnel: {ngrok_error}")
            return
        
        server_url, attendance_url = start_server(
            host="0.0.0.0",  # Allow connections from any IP address
            port=5000,
            session_code=session_code,
            callback=lambda name, id, validate_only=False: mark_student_attendance(
                excel_handler, lecture_date, name, id, validate_only
            )
        )

        # Generate QR Code with the public URL from ngrok
        qr_attendance_url = f"{public_url}/attendance?session={session_code}"
        qr_output_path = QR_OUTPUT_FILENAME
        _, _, _ = generate_lecture_qr(lecture_name, lecture_date, qr_attendance_url, qr_output_path)
        
        print(f"\nWeb server started at: {server_url}")
        print(f"Attendance URL: {qr_attendance_url}")
        print(f"QR code saved to: {os.path.abspath(qr_output_path)}")
        
        try:
            webbrowser.open(f"file://{os.path.abspath(qr_output_path)}")
            print("\nOpened QR code image - show this to students to scan")
        except:
            print(f"\nCould not open QR code image automatically. Please open it manually from: {os.path.abspath(qr_output_path)}")
        
        timer_thread = threading.Thread(
            target=attendance_timer,
            args=(excel_handler, lecture_date, custom_attendance_duration)
        )
        timer_thread.daemon = True
        timer_thread.start()
        
        print("\nWaiting for students to scan the QR code and register attendance...")
        print(f"Session will automatically end after {ATTENDANCE_DURATION//60} minute(s)")
        
        try:
            timer_thread.join()
        except KeyboardInterrupt:
            print("\nAttendance session manually stopped")
            excel_handler.mark_all_absent(lecture_date)
            excel_handler.save_file()
            print("Data saved. Exiting program...")
            os._exit(0)
        
    except Exception as e:
        print(f"Error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            ngrok.disconnect()
        except:
            pass
        
        print("\nThank you for using the QR Attendance System. Goodbye!")

if __name__ == "__main__":
    main()