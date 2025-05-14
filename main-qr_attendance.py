import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import threading
import time
import sys
import io
import pandas as pd
import webbrowser
import socket
import subprocess
import random
import urllib.request
import tempfile
import base64
import shutil

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    from qr_attendance.excel_handler import ExcelHandler
    from qr_attendance.generate_qr import generate_lecture_qr
    from qr_attendance.web_server import start_server
    from qr_attendance.config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME
except ImportError:
    from excel_handler import ExcelHandler
    from generate_qr import generate_lecture_qr
    from web_server import start_server
    from config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return socket.gethostbyname(socket.gethostname())

def create_simple_tunnel(port=5000):
    print("Setting up external access tunnel...")
    
    temp_dir = os.path.join(tempfile.gettempdir(), "attendance_system")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    tunnel_path = os.path.join(temp_dir, "expose.exe")
    
    if not os.path.exists(tunnel_path):
        try:
            print("Downloading required components...")
            
            binary_url = "https://github.com/beyondcode/expose/releases/latest/download/expose-windows.zip"
            zip_path = os.path.join(temp_dir, "expose.zip")
            
            urllib.request.urlretrieve(binary_url, zip_path)
            
            import zipfile
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            os.remove(zip_path)
            
            print("Components downloaded successfully")
        except Exception as e:
            print(f"Error downloading components: {e}")
            return None, None
    
    config_path = os.path.join(temp_dir, "expose.json")
    with open(config_path, 'w') as f:
        f.write('{"default-server": "expose.dev", "auth": ""}')
    
    try:
        print("Starting tunnel connection...")
        cmd = f"{tunnel_path} share http://localhost:{port} --server=sharedwithexpose.com"
        
        process = subprocess.Popen(
            cmd,
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True
        )
        
        url = None
        for i in range(30):
            line = process.stdout.readline()
            if "https://" in line:
                url = line.strip().split()[-1]
                if url.startswith("https://"):
                    print(f"Tunnel created: {url}")
                    return url, process
            time.sleep(0.5)
        
        process.kill()
        print("Failed to create tunnel - timeout")
        return None, None
        
    except Exception as e:
        print(f"Error creating tunnel: {e}")
        return None, None

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    
    return file_path if file_path else None

def attendance_timer(excel_handler, lecture_date, attendance_duration=ATTENDANCE_DURATION, tunnel_process=None):
    print(f"\nStarting attendance session for {attendance_duration//60} minute(s)...")
    time.sleep(attendance_duration)
    
    print("\n=== Attendance time expired ===")
    absent_count = excel_handler.mark_all_absent(lecture_date)
    print(f"{absent_count} students marked as absent")
    
    success, message = excel_handler.save_file()
    print(message)
    
    print("\nAttendance session closed and data saved")
    
    try:
        output_files = [os.path.abspath(OUTPUT_FILENAME), os.path.abspath(QR_OUTPUT_FILENAME)]
        for file_path in output_files:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"تم حذف الملف: {file_path}")
    except Exception as e:
        print(f"خطأ أثناء حذف الملفات: {e}")
    
    if tunnel_process:
        try:
            tunnel_process.terminate()
            print("Tunnel connection closed")
        except:
            pass
    
    print("\nProgram will close in 3 seconds...")
    time.sleep(3)
    os._exit(0)

def mark_student_attendance(excel_handler, lecture_date, student_name, student_id, validate_only=False):
    full_name_col = COLUMN_NAMES["full_name"]
    student_id_col = COLUMN_NAMES["student_id"]
    date_col = COLUMN_NAMES["date"]
    
    student_mask = (
        (excel_handler.df[student_id_col].astype(str).str.strip() == str(student_id).strip()) &
        (excel_handler.df[date_col].dt.date == lecture_date.date())
    )
    
    if not any(student_mask):
        return False, "خطأ: الطالب غير مسجل في محاضرة اليوم"
    
    if validate_only:
        return True, "الطالب موجود في قائمة المحاضرة"
    
    success, message = excel_handler.mark_attendance(student_name, student_id, lecture_date)
    print(f"Web submission: {student_name} ({student_id}) - {message}")
    return success, message

def reset_previous_attendance_data(excel_handler, lecture_date):
    count = excel_handler.reset_attendance_for_date(lecture_date)
    print(f"Reset attendance data for {count} students for date {lecture_date.strftime('%Y-%m-%d')}")
    
    success, message = excel_handler.save_file()
    if success:
        print("File saved successfully after resetting attendance data")
    else:
        print(f"Warning: {message}")
    
    return count

def create_security_dir():
    security_dir = resource_path('security')
    if not os.path.exists(security_dir):
        os.makedirs(security_dir)
    
    init_file = os.path.join(security_dir, '__init__.py')
    if not os.path.exists(init_file):
        with open(init_file, 'w') as f:
            f.write('# Security module initialization')

def main():
    tunnel_process = None
    
    try:
        print("=== QR Attendance System (Web Version) ===")
        
        create_security_dir()
        
        excel_file = select_excel_file()
        if not excel_file:
            print("No file selected. Exiting program.")
            return
        
        excel_handler = ExcelHandler(excel_file)
        excel_handler.load_file()
        
        date_col = COLUMN_NAMES["date"]
        if date_col in excel_handler.df.columns:
            sample_dates = excel_handler.df[date_col].head(3).tolist()
            print(f"Sample dates before conversion: {sample_dates}")
        
        excel_handler.convert_date_column()
        
        if date_col in excel_handler.df.columns:
            try:
                sample_dates = excel_handler.df[date_col].head(3).dt.strftime('%Y-%m-%d').tolist()
                print(f"Sample dates after conversion: {sample_dates}")
            except:
                print("Could not format converted dates")
        
        today_lectures = excel_handler.check_lecture_today()
        if today_lectures is None or today_lectures.empty:
            print("No lecture scheduled for today.")
            today = pd.Timestamp.now().date()
            print(f"Today's date is: {today}")
            
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("لا توجد محاضرة", "لا توجد محاضرة مجدولة لليوم الحالي.")
            
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
        
        print("\n=== Resetting previous attendance data ===")
        reset_previous_attendance_data(excel_handler, lecture_date)
        
        root = tk.Tk()
        root.withdraw()
        
        default_duration = ATTENDANCE_DURATION // 60
        
        duration_minutes = simpledialog.askinteger(
            "مدة التحضير", 
            "أدخل مدة التحضير بالدقائق:",
            initialvalue=default_duration,
            minvalue=1,
            maxvalue=120
        )
        
        if duration_minutes is None:
            duration_minutes = default_duration
            print(f"تم استخدام مدة التحضير الافتراضية: {duration_minutes} دقيقة")
        else:
            print(f"تم تحديد مدة التحضير: {duration_minutes} دقيقة")
        
        custom_attendance_duration = duration_minutes * 60
        
        session_code = str(int(time.time()))[-8:]
        port = 5000
        
        server_url, _ = start_server(
            host="0.0.0.0", 
            port=port,
            session_code=session_code,
            callback=lambda name, id, validate_only=False: mark_student_attendance(
                excel_handler, lecture_date, name, id, validate_only
            )
        )
        
        public_url, tunnel_process = create_simple_tunnel(port)
        
        if not public_url:
            public_url = f"http://{get_local_ip()}:{port}"
            
            msg = "يرجى التأكد من أن جميع الطلاب متصلين بنفس الشبكة المحلية."
            messagebox.showinfo("ملاحظة", msg)
        
        qr_attendance_url = f"{public_url}/attendance?session={session_code}"
        qr_output_path = QR_OUTPUT_FILENAME
        _, _, _ = generate_lecture_qr(lecture_name, lecture_date, qr_attendance_url, qr_output_path)
        
        print(f"\nWeb server started at: {server_url}")
        print(f"Public attendance URL: {qr_attendance_url}")
        print(f"QR code saved to: {os.path.abspath(qr_output_path)}")
        
        try:
            webbrowser.open(f"file://{os.path.abspath(qr_output_path)}")
            print("\nOpened QR code image - show this to students to scan")
        except:
            print(f"\nCould not open QR code image automatically. Please open it manually from: {os.path.abspath(qr_output_path)}")
        
        timer_thread = threading.Thread(
            target=attendance_timer,
            args=(excel_handler, lecture_date, custom_attendance_duration, tunnel_process)
        )
        timer_thread.daemon = True
        timer_thread.start()
        
        print("\nWaiting for students to scan the QR code and register attendance...")
        print(f"Session will automatically end after {duration_minutes} minute(s)")
        print("\nNOTE: Student attendance will be rejected if there is a fingerprint mismatch")
        
        control_root = tk.Tk()
        control_root.title("QR Attendance Control")
        control_root.geometry("400x200")
        control_root.resizable(False, False)
        
        def end_session():
            if messagebox.askyesno("تأكيد", "هل أنت متأكد من رغبتك في إنهاء جلسة التحضير؟"):
                excel_handler.mark_all_absent(lecture_date)
                excel_handler.save_file()
                
                if tunnel_process:
                    try:
                        tunnel_process.terminate()
                    except:
                        pass
                
                control_root.destroy()
                os._exit(0)
        
        label = tk.Label(control_root, text=f"جلسة التحضير: {lecture_name}", font=("Arial", 14))
        label.pack(pady=20)
        
        fingerprint_label = tk.Label(control_root, text="التحقق من البصمة مفعّل", font=("Arial", 10), fg="red")
        fingerprint_label.pack(pady=5)
        
        time_label = tk.Label(control_root, text=f"الوقت المتبقي: {duration_minutes} دقيقة", font=("Arial", 12))
        time_label.pack(pady=10)
        
        end_button = tk.Button(control_root, text="إنهاء التحضير", command=end_session, bg="#ff5555", fg="white", font=("Arial", 12))
        end_button.pack(pady=20)
        
        remaining_seconds = custom_attendance_duration
        
        def update_timer():
            nonlocal remaining_seconds
            if remaining_seconds > 0:
                remaining_seconds -= 1
                minutes = remaining_seconds // 60
                seconds = remaining_seconds % 60
                time_label.config(text=f"الوقت المتبقي: {minutes}:{seconds:02d}")
                control_root.after(1000, update_timer)
            else:
                time_label.config(text="انتهت مدة التحضير")
        
        update_timer()
        control_root.protocol("WM_DELETE_WINDOW", end_session)
        
        control_root.mainloop()
        
    except Exception as e:
        print(f"Error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if tunnel_process:
            try:
                tunnel_process.terminate()
                print("Tunnel connection closed")
            except:
                pass
                
        print("\nThank you for using the QR Attendance System. Goodbye!")

if __name__ == "__main__":
    main()