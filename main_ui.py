import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import subprocess
import sys
import os
import threading
import time
import io
import pandas as pd
import webbrowser
from pathlib import Path
import socket
import tempfile
import random
import urllib.request

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
try:
    from qr_attendance.excel_handler import ExcelHandler
    from qr_attendance.generate_qr import generate_lecture_qr
    from qr_attendance.web_server import start_server
    from qr_attendance.config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME
except ImportError:
    try:
        from excel_handler import ExcelHandler
        from generate_qr import generate_lecture_qr
        from web_server import start_server
        from config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME
    except ImportError:
        pass

class ThemeManager:
    PRIMARY = "#1976d2"
    PRIMARY_DARK = "#004ba0" 
    PRIMARY_LIGHT = "#63a4ff"
    SECONDARY = "#283593"
    BG_LIGHT = "#212121"
    BG_DARK = "#f8f9fa"
    
    TEXT_LIGHT = "#212121"
    TEXT_GRAY_LIGHT = "#424242"  
    TEXT_DARK = "#ffffff"
    TEXT_DARK_GRAY = "#e0e0e0"  
    TEXT_ACCENT = "#90caf9"
    TEXT_SECONDARY = "#ce93d8"
    
    ACCENT = "#ff3d00"
    SUCCESS = "#00c853"
    WARNING = "#ffd600"
    ERROR = "#d50000"
    INFO = "#00b0ff"
    
    FONT_LARGE = ("Segoe UI", 22, "bold")
    FONT_MEDIUM = ("Segoe UI", 14, "bold")
    FONT_SMALL = ("Segoe UI", 12)
    FONT_MINI = ("Segoe UI", 10)
    
    @classmethod
    def setup_styles(cls, style):
        style.configure(
            'Primary.TButton',
            font=cls.FONT_MEDIUM,
            background=cls.PRIMARY,
            foreground=cls.TEXT_LIGHT,
            padding=(20, 15),
            borderwidth=0,
        )
        style.map(
            'Primary.TButton',
            background=[('active', cls.PRIMARY_LIGHT)],
            foreground=[('active', cls.TEXT_LIGHT)]
        )
        
        style.configure(
            'Graduation.TButton',
            font=("Segoe UI", 16, "bold"),
            background=cls.PRIMARY,
            foreground=cls.TEXT_LIGHT,
            padding=(20, 15),
            borderwidth=0,
        )
        style.map(
            'Graduation.TButton',
            background=[('active', cls.PRIMARY_LIGHT)],
            foreground=[('active', cls.TEXT_LIGHT)]
        )
        
        style.configure(
            'Sidebar.TButton',
            font=cls.FONT_SMALL,
            background=cls.BG_DARK,
            foreground=cls.TEXT_LIGHT,
            padding=(15, 12),
            borderwidth=0,
        )
        style.map(
            'Sidebar.TButton',
            background=[('active', cls.SECONDARY)],
            foreground=[('active', cls.TEXT_LIGHT)]
        )
        
        style.configure('Dark.TFrame', background=cls.BG_DARK)
        style.configure('Light.TFrame', background=cls.BG_LIGHT)
        
        style.configure(
            'Dark.TLabel',
            background=cls.BG_DARK,
            foreground=cls.TEXT_LIGHT,
            font=cls.FONT_SMALL
        )
        style.configure(
            'Light.TLabel',
            background=cls.BG_LIGHT,
            foreground=cls.TEXT_DARK,
            font=cls.FONT_SMALL
        )
        style.configure(
            'Title.TLabel',
            background=cls.BG_DARK,
            foreground=cls.TEXT_LIGHT,
            font=cls.FONT_LARGE
        )
        style.configure(
            'Heading.TLabel',
            background=cls.BG_LIGHT,
            foreground=cls.TEXT_ACCENT,
            font=cls.FONT_LARGE
        )
        
        style.configure(
            'Status.TLabel',
            background=cls.BG_LIGHT,
            foreground=cls.TEXT_DARK,
            font=cls.FONT_MINI
        )
        
        style.configure(
            'Version.TLabel',
            background=cls.BG_DARK,
            foreground=cls.TEXT_LIGHT,
            font=cls.FONT_MINI
        )
        
        style.configure('TSeparator', background=cls.PRIMARY_LIGHT)

class MyTVTCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MyTVTC")
        
        self.root.minsize(1100, 750)
        self.root.configure(bg=ThemeManager.BG_LIGHT)
        
        self.style = ttk.Style()
        ThemeManager.setup_styles(self.style)
        
        self.main_container = ttk.Frame(self.root, style='Light.TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        self.excel_handler = None
        self.lecture_date = None
        self.timer_thread = None
        self.tunnel_process = None
        self.qr_window = None
        
        self.create_sidebar()
        self.create_content_area()
        
        self.setup_app_icon()
        self.setup_responsive_layout()
        self.center_window(1100, 750)
        
    def create_sidebar(self):
        self.sidebar = ttk.Frame(self.main_container, style='Dark.TFrame', width=250)
        self.sidebar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        
        app_title = ttk.Label(
            self.sidebar, 
            text="MyTVTC", 
            style='Title.TLabel'
        )
        app_title.pack(pady=(50, 30))
        
        separator = ttk.Separator(self.sidebar, orient='horizontal')
        separator.pack(fill=tk.X, padx=25, pady=20)
        
        self.developer_btn = ttk.Button(
            self.sidebar,
            text="معلومات المطور",
            style='Sidebar.TButton',
            command=self.show_developer_info
        )
        self.developer_btn.pack(fill=tk.X, pady=8, padx=20)
        
        version_frame = ttk.Frame(self.sidebar, style='Dark.TFrame')
        version_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=25)
        
        version_label = ttk.Label(
            version_frame,
            text="الإصدار 2.1.0",
            style='Version.TLabel'
        )
        version_label.pack(pady=5, anchor=tk.CENTER)
        
    def create_content_area(self):
        self.content = ttk.Frame(self.main_container, style='Light.TFrame')
        self.content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        header_frame = ttk.Frame(self.content, style='Light.TFrame')
        header_frame.pack(fill=tk.X, pady=(60, 50))
        
        heading = ttk.Label(
            header_frame,
            text="نظام التحضير الإلكتروني",
            style='Heading.TLabel'
        )
        heading.pack(anchor=tk.CENTER)
        
        btn_frame = ttk.Frame(self.content, style='Light.TFrame')
        btn_frame.pack(pady=40, padx=140)
        
        self.btn_scraping = ttk.Button(
            btn_frame,
            text="تصدير قالب التحضير",
            style='Primary.TButton',
            command=lambda: self.run_script("main-scraping.py")
        )
        self.btn_scraping.pack(pady=25, fill=tk.X)
        
        self.btn_qr_attendance = ttk.Button(
            btn_frame,
            text="بدء التحضير",
            style='Primary.TButton',
            command=lambda: self.run_script("main-qr_attendance.py")
        )
        self.btn_qr_attendance.pack(pady=25, fill=tk.X)
        
        self.btn_ext_main = ttk.Button(
            btn_frame,
            text="رفع التحضير في الموقع",
            style='Primary.TButton',
            command=lambda: self.run_script("main-finished.py")
        )
        self.btn_ext_main.pack(pady=25, fill=tk.X)
        
        self.btn_grad = ttk.Button(
            btn_frame,
            text="نظام رصد الدرجات",
            style='Graduation.TButton',
            command=lambda: self.run_script("main-Grad.py")
        )
        self.btn_grad.pack(pady=25, fill=tk.X)
        
        self.status_frame = ttk.Frame(self.content, style='Light.TFrame')
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=30, pady=20)
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="جاهز",
            style='Status.TLabel'
        )
        self.status_label.pack(side=tk.LEFT, padx=20)
    
    def run_script(self, script_name):
        try:
            self.update_status(f"جاري تشغيل {script_name}...")
            
            button_map = {
                "main-scraping.py": (self.btn_scraping, "تصدير قالب التحضير"),
                "main-qr_attendance.py": (self.btn_qr_attendance, "بدء التحضير"),
                "main-finished.py": (self.btn_ext_main, "رفع التحضير في الموقع"),
                "main-Grad.py": (self.btn_grad, "نظام رصد الدرجات")
            }
            
            current_button, original_text = button_map[script_name]
            current_button.config(text="جاري التنفيذ...")
            self.root.update()
            
            if script_name == "main-qr_attendance.py":
                self.start_qr_attendance()
                current_button.config(text=original_text)
                return
            
            script_path = Path(script_name)
            if not script_path.exists():
                messagebox.showwarning(
                    "تحذير",
                    f"الملف {script_name} غير موجود في المجلد الحالي."
                )
                current_button.config(text=original_text)
                self.update_status("جاهز")
                return
            
            process = subprocess.Popen([sys.executable, script_name])
            
            self.root.after(2000, lambda: current_button.config(text=original_text))
            self.root.after(2500, lambda: self.update_status("تم التنفيذ بنجاح"))
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تشغيل البرنامج:\n{str(e)}")
            if 'current_button' in locals() and 'original_text' in locals():
                current_button.config(text=original_text)
            self.update_status("حدث خطأ في التنفيذ")
    
    def update_status(self, message):
        self.status_label.config(text=message)
    
    def get_local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return socket.gethostbyname(socket.gethostname())
    
    def create_tunnel(self, port=5000):
        self.update_status("جاري إنشاء اتصال للوصول من خارج الشبكة...")
        
        temp_dir = os.path.join(tempfile.gettempdir(), "attendance_tunnel")
        os.makedirs(temp_dir, exist_ok=True)
        
        local_ip = self.get_local_ip()
        
        try:
            tunnel_cmd = 'ssh -o StrictHostKeyChecking=no -R 80:localhost:5000 localhost.run'
            process = subprocess.Popen(
                tunnel_cmd, 
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            import re
            for i in range(20):
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                
                match = re.search(r'https?://[a-zA-Z0-9\-]+\.localhost\.run', line)
                if match:
                    public_url = match.group(0)
                    self.update_status(f"تم إنشاء اتصال: {public_url}")
                    return public_url, process
                
                time.sleep(0.5)
            
            process.kill()
            
        except Exception as e:
            print(f"فشل في الاتصال باستخدام localhost.run: {e}")
        
        try:
            tunnel_cmd = 'ssh -o StrictHostKeyChecking=no -R 80:localhost:5000 serveo.net'
            process = subprocess.Popen(
                tunnel_cmd, 
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            import re
            for i in range(20):
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                
                match = re.search(r'https?://[a-zA-Z0-9\-\.]+\.serveo\.net', line)
                if match:
                    public_url = match.group(0)
                    self.update_status(f"تم إنشاء اتصال: {public_url}")
                    return public_url, process
                
                time.sleep(0.5)
            
            process.kill()
            
        except Exception as e:
            print(f"فشل في الاتصال باستخدام serveo.net: {e}")
        
        self.update_status("تعذر إنشاء اتصال خارجي، سيتم استخدام الشبكة المحلية")
        return f"http://{local_ip}:{port}", None
    
    def start_qr_attendance(self):
        try:
            self.update_status("جاري بدء نظام التحضير...")
            
            excel_file = filedialog.askopenfilename(
                title="اختر ملف التحضير",
                filetypes=[("ملفات إكسل", "*.xlsx;*.xls")]
            )
            
            if not excel_file:
                self.update_status("تم إلغاء العملية")
                return
            
            self.update_status("جاري تحليل ملف التحضير...")
            
            self.excel_handler = ExcelHandler(excel_file)
            self.excel_handler.load_file()
            
            date_col = COLUMN_NAMES["date"]
            self.excel_handler.convert_date_column()
            
            today_lectures = self.excel_handler.check_lecture_today()
            if today_lectures is None or today_lectures.empty:
                self.show_styled_message(
                    "لا توجد محاضرة",
                    "لا توجد محاضرة مجدولة لليوم الحالي.",
                    "warning"
                )
                self.update_status("لا توجد محاضرة اليوم")
                return
            
            lecture_info = today_lectures.iloc[0]
            
            lecture_name_col = COLUMN_NAMES["lecture_name"]
            section_col = COLUMN_NAMES["section"]
            date_col = COLUMN_NAMES["date"]
            
            lecture_name = "محاضرة غير معروفة"
            if lecture_name_col in lecture_info.index and not pd.isnull(lecture_info[lecture_name_col]):
                lecture_name = lecture_info[lecture_name_col]
            elif section_col in lecture_info.index and not pd.isnull(lecture_info[section_col]):
                lecture_name = lecture_info[section_col]
                
            self.lecture_date = lecture_info[date_col]
            
            self.update_status(f"تم العثور على محاضرة: {lecture_name}")
            
            self.excel_handler.reset_attendance_for_date(self.lecture_date)
            
            duration_minutes = self.show_duration_dialog()
            if duration_minutes is None:
                self.update_status("تم إلغاء العملية")
                return
                
            custom_attendance_duration = duration_minutes * 60
            
            session_code = str(int(time.time()))[-8:]
            port = 5000
            
            def mark_attendance_with_fingerprint(student_name, student_id, validate_only=False):
                if not self.excel_handler or not self.lecture_date:
                    return False, "خطأ في النظام: لم يتم تهيئة نظام التحضير بشكل صحيح"
                
                result = self.excel_handler.mark_attendance(student_name, student_id, self.lecture_date)
                
                if not validate_only and result[0] and self.qr_window and self.qr_window.winfo_exists():
                    present_count = len(self.excel_handler.present_students)
                    self.present_count_var.set(str(present_count))
                    
                return result
            
            server_url, _ = start_server(
                host="0.0.0.0",
                port=port,
                session_code=session_code,
                callback=mark_attendance_with_fingerprint
            )
            
            public_url, self.tunnel_process = self.create_tunnel(port)
            
            qr_attendance_url = f"{public_url}/attendance?session={session_code}"
            qr_output_path = QR_OUTPUT_FILENAME
            _, _, _ = generate_lecture_qr(lecture_name, self.lecture_date, qr_attendance_url, qr_output_path)
            
            self.update_status("تم إنشاء رمز QR بنجاح")
            
            self.show_qr_window(qr_output_path, lecture_name, duration_minutes)
            
            self.timer_thread = threading.Thread(
                target=self.attendance_timer_thread,
                args=(custom_attendance_duration,)
            )
            self.timer_thread.daemon = True
            self.timer_thread.start()
            
        except Exception as e:
            self.show_styled_message(
                "خطأ",
                f"حدث خطأ أثناء بدء نظام التحضير:\n{str(e)}",
                "error"
            )
            self.update_status("فشل في بدء نظام التحضير")
            self.cleanup_resources()
    
    def show_duration_dialog(self):
        duration_dialog = tk.Toplevel(self.root)
        duration_dialog.title("تحديد مدة التحضير")
        duration_dialog.configure(bg=ThemeManager.BG_LIGHT)
        duration_dialog.resizable(False, False)
        
        self.center_window_on_parent(duration_dialog, 400, 250)
        
        duration_dialog.transient(self.root)
        duration_dialog.grab_set()
        
        main_frame = ttk.Frame(duration_dialog, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(
            main_frame,
            text="تحديد مدة التحضير",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        title_label.pack(pady=(0, 20))
        
        desc_label = ttk.Label(
            main_frame,
            text="الرجاء تحديد مدة التحضير بالدقائق:",
            style='Light.TLabel'
        )
        desc_label.pack(pady=(0, 15))
        
        duration_var = tk.IntVar(value=ATTENDANCE_DURATION // 60)
        
        entry_frame = ttk.Frame(main_frame, style='Light.TFrame')
        entry_frame.pack(pady=(0, 20), fill=tk.X)
        
        duration_entry = ttk.Entry(
            entry_frame,
            textvariable=duration_var,
            width=10,
            font=("Segoe UI", 14)
        )
        duration_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        minutes_label = ttk.Label(
            entry_frame,
            text="دقيقة",
            style='Light.TLabel'
        )
        minutes_label.pack(side=tk.LEFT)
        
        button_frame = ttk.Frame(main_frame, style='Light.TFrame')
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        result = [None]
        
        def confirm():
            try:
                value = duration_var.get()
                if value < 1:
                    messagebox.showwarning("تحذير", "يجب أن تكون المدة دقيقة واحدة على الأقل")
                    return
                if value > 120:
                    messagebox.showwarning("تحذير", "الحد الأقصى للمدة هو 120 دقيقة")
                    return
                result[0] = value
                duration_dialog.destroy()
            except:
                messagebox.showwarning("تحذير", "الرجاء إدخال قيمة صحيحة")
        
        def cancel():
            result[0] = None
            duration_dialog.destroy()
        
        confirm_button = ttk.Button(
            button_frame,
            text="تأكيد",
            style='Primary.TButton',
            command=confirm
        )
        confirm_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        cancel_button = ttk.Button(
            button_frame,
            text="إلغاء",
            command=cancel
        )
        cancel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        duration_dialog.bind('<Return>', lambda event: confirm())
        
        self.root.wait_window(duration_dialog)
        
        return result[0]
    
    def show_qr_window(self, qr_path, lecture_name, duration_minutes):
        if self.qr_window and self.qr_window.winfo_exists():
            self.qr_window.destroy()
        
        self.qr_window = tk.Toplevel(self.root)
        self.qr_window.title("رمز QR للتحضير")
        self.qr_window.configure(bg=ThemeManager.BG_LIGHT)
        self.qr_window.minsize(650, 750)
        
        self.center_window_on_parent(self.qr_window, 650, 750)
        
        main_frame = ttk.Frame(self.qr_window, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(
            main_frame,
            text=f"تحضير: {lecture_name}",
            font=("Segoe UI", 18, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT,
            wraplength=590
        )
        title_label.pack(pady=(0, 10))
        
        info_frame = ttk.Frame(main_frame, style='Light.TFrame')
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        desc_label = ttk.Label(
            info_frame,
            text="اطلب من الطلاب مسح رمز QR التالي لتسجيل الحضور",
            font=("Segoe UI", 12),
            style='Light.TLabel',
            wraplength=590
        )
        desc_label.pack(pady=(0, 5))
        
        network_label = ttk.Label(
            info_frame,
            text="يمكن للطلاب استخدام أي شبكة إنترنت (واي فاي أو بيانات)",
            foreground=ThemeManager.SUCCESS,
            background=ThemeManager.BG_LIGHT,
            font=("Segoe UI", 12)
        )
        network_label.pack(pady=(0, 5))
        
        fingerprint_label = ttk.Label(
            info_frame,
            text="سيتم التحقق من بصمة الجهاز عند التسجيل، وسيتم رفض التسجيل في حالة وجود تعارض",
            foreground=ThemeManager.ERROR,
            background=ThemeManager.BG_LIGHT,
            font=("Segoe UI", 11)
        )
        fingerprint_label.pack(pady=(0, 10))
        
        try:
            from PIL import Image, ImageTk
            
            qr_image = Image.open(qr_path)
            qr_image = qr_image.resize((400, 400), Image.LANCZOS)
            qr_photo = ImageTk.PhotoImage(qr_image)
            
            qr_label = tk.Label(main_frame, image=qr_photo, bg=ThemeManager.BG_LIGHT)
            qr_label.image = qr_photo
            qr_label.pack(pady=10)
        except Exception as e:
            qr_error_label = ttk.Label(
                main_frame,
                text=f"تعذر عرض رمز QR: {str(e)}",
                style='Light.TLabel'
            )
            qr_error_label.pack(pady=10)
            
            webbrowser.open(f"file://{os.path.abspath(qr_path)}")
        
        present_frame = ttk.Frame(main_frame, style='Light.TFrame')
        present_frame.pack(pady=(10, 5), fill=tk.X)
        
        present_label = ttk.Label(
            present_frame,
            text="الطلاب الحاضرون:",
            font=("Segoe UI", 14, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        present_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.present_count_var = tk.StringVar(value="0")
        
        count_label = ttk.Label(
            present_frame,
            textvariable=self.present_count_var,
            font=("Segoe UI", 14, "bold"),
            foreground=ThemeManager.SUCCESS,
            background=ThemeManager.BG_LIGHT
        )
        count_label.pack(side=tk.LEFT)
        
        time_frame = ttk.Frame(main_frame, style='Light.TFrame')
        time_frame.pack(pady=10, fill=tk.X)
        
        time_label = ttk.Label(
            time_frame,
            text=f"مدة التحضير: {duration_minutes} دقيقة",
            font=("Segoe UI", 14, "bold"),
            foreground=ThemeManager.TEXT_ACCENT,
            background=ThemeManager.BG_LIGHT
        )
        time_label.pack()
        
        self.remaining_time_var = tk.StringVar(value="جاري بدء التحضير...")
        
        remaining_label = ttk.Label(
            time_frame,
            textvariable=self.remaining_time_var,
            font=("Segoe UI", 12),
            style='Light.TLabel'
        )
        remaining_label.pack(pady=(5, 0))
        
        close_button = ttk.Button(
            main_frame,
            text="إنهاء التحضير",
            style='Primary.TButton',
            command=self.end_attendance_session
        )
        close_button.pack(pady=(20, 0), fill=tk.X)
        
        self.qr_window.protocol("WM_DELETE_WINDOW", self.end_attendance_session)
    
    def attendance_timer_thread(self, duration_seconds):
        start_time = time.time()
        end_time = start_time + duration_seconds
        
        try:
            while time.time() < end_time:
                remaining = end_time - time.time()
                minutes = int(remaining // 60)
                seconds = int(remaining % 60)
                
                if self.qr_window and self.qr_window.winfo_exists():
                    self.remaining_time_var.set(f"الوقت المتبقي: {minutes:02d}:{seconds:02d}")
                    
                    if self.excel_handler:
                        present_count = len(self.excel_handler.present_students)
                        self.present_count_var.set(str(present_count))
                
                time.sleep(1)
            
            self.complete_attendance_session()
            
        except Exception as e:
            print(f"خطأ في مؤقت الحضور: {e}")
    
    def mark_student_attendance(self, student_name, student_id, validate_only=False):
        if not self.excel_handler or not self.lecture_date:
            return False, "خطأ في النظام: لم يتم تهيئة نظام التحضير بشكل صحيح"
        
        result = self.excel_handler.mark_attendance(student_name, student_id, self.lecture_date)
        
        if not validate_only and result[0] and self.qr_window and self.qr_window.winfo_exists():
            present_count = len(self.excel_handler.present_students)
            self.present_count_var.set(str(present_count))
            
        return result
    
    def end_attendance_session(self):
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من رغبتك في إنهاء جلسة التحضير؟"):
            self.complete_attendance_session()
    
    def complete_attendance_session(self):
        try:
            if self.excel_handler and self.lecture_date:
                absent_count = self.excel_handler.mark_all_absent(self.lecture_date)
                
                success, message = self.excel_handler.save_file()
                
                self.show_completion_message(absent_count)
                
                self.cleanup_files()
            
            self.cleanup_resources()
            
            self.update_status("تم إكمال جلسة التحضير بنجاح")
            
        except Exception as e:
            self.show_styled_message(
                "خطأ",
                f"حدث خطأ أثناء إكمال جلسة التحضير:\n{str(e)}",
                "error"
            )
            self.update_status("حدث خطأ أثناء إكمال جلسة التحضير")
            self.cleanup_resources()
    
    def show_completion_message(self, absent_count):
        completion_dialog = tk.Toplevel(self.root)
        completion_dialog.title("اكتمال التحضير")
        completion_dialog.configure(bg=ThemeManager.BG_LIGHT)
        completion_dialog.resizable(False, False)
        
        self.center_window_on_parent(completion_dialog, 450, 350)
        
        completion_dialog.transient(self.root)
        completion_dialog.grab_set()
        
        main_frame = ttk.Frame(completion_dialog, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        success_label = ttk.Label(
            main_frame,
            text="✓",
            font=("Segoe UI", 48),
            foreground=ThemeManager.SUCCESS,
            background=ThemeManager.BG_LIGHT
        )
        success_label.pack(pady=(0, 10))
        
        title_label = ttk.Label(
            main_frame,
            text="تم اكتمال التحضير بنجاح",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        title_label.pack(pady=(0, 20))
        
        info_frame = ttk.Frame(main_frame, style='Light.TFrame')
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        present_count = len(self.excel_handler.present_students) if self.excel_handler else 0
        
        present_text = f"تم تسجيل {present_count} طالب كحاضر"
        present_label = ttk.Label(
            info_frame,
            text=present_text,
            foreground=ThemeManager.SUCCESS,
            background=ThemeManager.BG_LIGHT,
            font=("Segoe UI", 12)
        )
        present_label.pack(pady=5)
        
        absent_text = f"تم تسجيل {absent_count} طالب كغائب"
        absent_label = ttk.Label(
            info_frame,
            text=absent_text,
            style='Light.TLabel'
        )
        absent_label.pack()
        
        save_label = ttk.Label(
            info_frame,
            text="تم حفظ بيانات التحضير في ملف الإكسل",
            style='Light.TLabel'
        )
        save_label.pack(pady=(5, 0))
        
        close_button = ttk.Button(
            main_frame,
            text="إغلاق",
            style='Primary.TButton',
            command=completion_dialog.destroy
        )
        close_button.pack(fill=tk.X, pady=(10, 0))
        
        completion_dialog.bind('<Return>', lambda event: completion_dialog.destroy())
    
    def cleanup_resources(self):
        if self.qr_window and self.qr_window.winfo_exists():
            self.qr_window.destroy()
            self.qr_window = None
        
        if self.tunnel_process:
            try:
                self.tunnel_process.terminate()
                print("تم إغلاق اتصال النفق")
            except:
                pass
            self.tunnel_process = None
        
        self.excel_handler = None
        self.lecture_date = None
        self.timer_thread = None
    
    def cleanup_files(self):
        try:
            output_files = [os.path.abspath(OUTPUT_FILENAME), os.path.abspath(QR_OUTPUT_FILENAME)]
            for file_path in output_files:
                if os.path.exists(file_path):
                    os.remove(file_path)
        except Exception as e:
            print(f"خطأ أثناء حذف الملفات: {e}")
    
    def show_styled_message(self, title, message, message_type="info"):
        if message_type == "error":
            messagebox.showerror(title, message)
        elif message_type == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)
    
    def show_developer_info(self):
        dev_window = tk.Toplevel(self.root)
        dev_window.title("معلومات المطور")
        dev_window.geometry("700x600")
        dev_window.configure(bg=ThemeManager.BG_LIGHT)
        dev_window.resizable(True, True)
        
        self.center_window_on_parent(dev_window, 700, 600)
        
        dev_window.transient(self.root)
        dev_window.grab_set()
        
        info_frame = ttk.Frame(dev_window, style='Light.TFrame')
        info_frame.pack(padx=40, pady=(40, 20), fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(
            info_frame,
            text="فريق التطوير",
            style='Heading.TLabel'
        )
        title_label.pack(pady=(15, 30))
        
        info_text = ttk.Label(
            info_frame,
            text="تم تطوير نظام التحضير الإلكتروني بواسطة:",
            style='Light.TLabel'
        )
        info_text.pack(pady=10)
        
        dev_name1 = ttk.Label(
            info_frame,
            text="م. سعد ياسر العلي",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.TEXT_ACCENT,
            background=ThemeManager.BG_LIGHT
        )
        dev_name1.pack(pady=5)
        
        dev_name2 = ttk.Label(
            info_frame,
            text="م. عمار الزهراني",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.TEXT_ACCENT,
            background=ThemeManager.BG_LIGHT
        )
        dev_name2.pack(pady=5)
        
        phone_label = ttk.Label(
            info_frame,
            text="للاستفسارات والدعم الفني: 0557000646",
            style='Light.TLabel'
        )
        phone_label.pack(pady=15)
        
        system_frame = ttk.Frame(info_frame, style='Light.TFrame')
        system_frame.pack(pady=20, fill=tk.X)
        
        system_info = ttk.Label(
            system_frame,
            text=f"MyTVTC - نظام التحضير الإلكتروني المتكامل\nالإصدار 2.1.0",
            style='Light.TLabel',
            justify=tk.CENTER
        )
        system_info.pack()
        
        button_frame = ttk.Frame(info_frame, style='Light.TFrame')
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=40)
        
        close_btn = ttk.Button(
            button_frame,
            text="إغلاق",
            style='Primary.TButton',
            command=dev_window.destroy
        )
        close_btn.pack(pady=10, fill=tk.X, padx=20)
    
    def setup_app_icon(self):
        try:
            icon_path = Path("app_icon.ico")
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
        except Exception:
            pass
    
    def setup_responsive_layout(self):
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(1, weight=1)
    
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def center_window_on_parent(self, window, width, height):
        parent_x = self.root.winfo_rootx()
        parent_y = self.root.winfo_rooty()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        
        x = parent_x + int((parent_width / 2) - (width / 2))
        y = parent_y + int((parent_height / 2) - (height / 2))
        
        window.geometry(f"{width}x{height}+{x}+{y}")

def main():
    root = tk.Tk()
    
    root.tk.call('encoding', 'system', 'utf-8')
    try:
        root.tk.call('package', 'require', 'Ttk')
        root.tk.call('ttk::style', 'configure', '.', '-font', 'TkDefaultFont')
    except:
        pass
    
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    app = MyTVTCApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()