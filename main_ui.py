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
from pyngrok import ngrok
from pyngrok import conf as ngrok_conf

# Import QR attendance modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
try:
    from qr_attendance.excel_handler import ExcelHandler
    from qr_attendance.generate_qr import generate_lecture_qr
    from qr_attendance.web_server import start_server
    from qr_attendance.config import COLUMN_NAMES, ATTENDANCE_DURATION, QR_OUTPUT_FILENAME, OUTPUT_FILENAME
except ImportError:
    # Fallback for direct imports
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
        
        self.root.minsize(950, 650)
        self.root.configure(bg=ThemeManager.BG_LIGHT)
        
        self.style = ttk.Style()
        ThemeManager.setup_styles(self.style)
        
        self.main_container = ttk.Frame(self.root, style='Light.TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # متغيرات لنظام التحضير
        self.excel_handler = None
        self.lecture_date = None
        self.timer_thread = None
        self.ngrok_tunnel = None
        self.qr_window = None
        
        self.create_sidebar()
        self.create_content_area()
        
        self.setup_app_icon()
        self.setup_responsive_layout()
        self.center_window(950, 650)
        
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
            text="الإصدار 2.0.0",
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
        btn_frame.pack(pady=30, padx=100)
        
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
            command=lambda: self.run_script("ext-main.py")
        )
        self.btn_ext_main.pack(pady=25, fill=tk.X)
        
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
                "ext-main.py": (self.btn_ext_main, "رفع التحضير في الموقع")
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
    
    def start_qr_attendance(self):
        """بدء نظام التحضير بالـ QR مباشرة من الواجهة الرئيسية"""
        try:
            self.update_status("جاري بدء نظام التحضير...")
            
            # Configure ngrok
            ngrok_conf.get_default().authtoken = '2uJbHU37cEkmp2othhPYSZjNY8N_5trhFGNwJKjswwSp89k6Y'
            
            # اختيار ملف الإكسل
            excel_file = filedialog.askopenfilename(
                title="اختر ملف التحضير",
                filetypes=[("ملفات إكسل", "*.xlsx;*.xls")]
            )
            
            if not excel_file:
                self.update_status("تم إلغاء العملية")
                return
            
            self.update_status("جاري تحليل ملف التحضير...")
            
            # تحميل ملف الإكسل
            self.excel_handler = ExcelHandler(excel_file)
            self.excel_handler.load_file()
            
            # تحويل عمود التاريخ
            date_col = COLUMN_NAMES["date"]
            self.excel_handler.convert_date_column()
            
            # التحقق من وجود محاضرة اليوم
            today_lectures = self.excel_handler.check_lecture_today()
            if today_lectures is None or today_lectures.empty:
                self.show_styled_message(
                    "لا توجد محاضرة",
                    "لا توجد محاضرة مجدولة لليوم الحالي.",
                    "warning"
                )
                self.update_status("لا توجد محاضرة اليوم")
                return
            
            # استخراج معلومات المحاضرة
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
            
            # إعادة تعيين بيانات الحضور السابقة
            self.excel_handler.reset_attendance_for_date(self.lecture_date)
            
            # طلب مدة التحضير من المستخدم
            duration_minutes = self.show_duration_dialog()
            if duration_minutes is None:
                self.update_status("تم إلغاء العملية")
                return
                
            # تحويل الدقائق إلى ثواني
            custom_attendance_duration = duration_minutes * 60
            
            # إنشاء رمز الجلسة
            session_code = str(int(time.time()))[-8:]
            
            # فتح نفق ngrok
            try:
                self.ngrok_tunnel = ngrok.connect(5000)
                public_url = self.ngrok_tunnel.public_url
            except Exception as ngrok_error:
                self.show_styled_message(
                    "خطأ في الاتصال",
                    f"فشل في إنشاء اتصال ngrok: {ngrok_error}",
                    "error"
                )
                self.update_status("فشل في بدء نظام التحضير")
                return
            
            # بدء خادم الويب
            server_url, attendance_url = start_server(
                host="0.0.0.0",
                port=5000,
                session_code=session_code,
                callback=lambda name, id, validate_only=False: self.mark_student_attendance(name, id, validate_only)
            )

            # إنشاء رمز QR
            qr_attendance_url = f"{public_url}/attendance?session={session_code}"
            qr_output_path = QR_OUTPUT_FILENAME
            _, _, _ = generate_lecture_qr(lecture_name, self.lecture_date, qr_attendance_url, qr_output_path)
            
            self.update_status("تم إنشاء رمز QR بنجاح")
            
            # عرض رمز QR في نافذة منفصلة
            self.show_qr_window(qr_output_path, lecture_name, duration_minutes)
            
            # بدء مؤقت الحضور
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
        """عرض نافذة لتحديد مدة التحضير بتصميم محسن"""
        duration_dialog = tk.Toplevel(self.root)
        duration_dialog.title("تحديد مدة التحضير")
        duration_dialog.configure(bg=ThemeManager.BG_LIGHT)
        duration_dialog.resizable(False, False)
        
        # جعل النافذة مركزية
        self.center_window_on_parent(duration_dialog, 400, 250)
        
        # جعل النافذة مودال (تمنع التفاعل مع النافذة الأم)
        duration_dialog.transient(self.root)
        duration_dialog.grab_set()
        
        # إطار رئيسي
        main_frame = ttk.Frame(duration_dialog, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        # عنوان
        title_label = ttk.Label(
            main_frame,
            text="تحديد مدة التحضير",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        title_label.pack(pady=(0, 20))
        
        # وصف
        desc_label = ttk.Label(
            main_frame,
            text="الرجاء تحديد مدة التحضير بالدقائق:",
            style='Light.TLabel'
        )
        desc_label.pack(pady=(0, 15))
        
        # متغير لتخزين القيمة
        duration_var = tk.IntVar(value=ATTENDANCE_DURATION // 60)
        
        # إطار للمدخل
        entry_frame = ttk.Frame(main_frame, style='Light.TFrame')
        entry_frame.pack(pady=(0, 20), fill=tk.X)
        
        # مدخل المدة
        duration_entry = ttk.Entry(
            entry_frame,
            textvariable=duration_var,
            width=10,
            font=("Segoe UI", 14)
        )
        duration_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # نص الدقائق
        minutes_label = ttk.Label(
            entry_frame,
            text="دقيقة",
            style='Light.TLabel'
        )
        minutes_label.pack(side=tk.LEFT)
        
        # إطار للأزرار
        button_frame = ttk.Frame(main_frame, style='Light.TFrame')
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # متغير للنتيجة
        result = [None]
        
        # دالة للتأكيد
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
        
        # دالة للإلغاء
        def cancel():
            result[0] = None
            duration_dialog.destroy()
        
        # زر التأكيد
        confirm_button = ttk.Button(
            button_frame,
            text="تأكيد",
            style='Primary.TButton',
            command=confirm
        )
        confirm_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # زر الإلغاء
        cancel_button = ttk.Button(
            button_frame,
            text="إلغاء",
            command=cancel
        )
        cancel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        # تعيين زر التأكيد كزر افتراضي
        duration_dialog.bind('<Return>', lambda event: confirm())
        
        # انتظار إغلاق النافذة
        self.root.wait_window(duration_dialog)
        
        return result[0]
    
    def show_qr_window(self, qr_path, lecture_name, duration_minutes):
        """عرض رمز QR في نافذة منفصلة"""
        if self.qr_window and self.qr_window.winfo_exists():
            self.qr_window.destroy()
        
        self.qr_window = tk.Toplevel(self.root)
        self.qr_window.title("رمز QR للتحضير")
        self.qr_window.configure(bg=ThemeManager.BG_LIGHT)
        self.qr_window.minsize(500, 600)
        
        # جعل النافذة مركزية
        self.center_window_on_parent(self.qr_window, 500, 600)
        
        # إطار رئيسي
        main_frame = ttk.Frame(self.qr_window, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        # عنوان
        title_label = ttk.Label(
            main_frame,
            text=f"تحضير: {lecture_name}",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        title_label.pack(pady=(0, 10))
        
        # وصف
        desc_label = ttk.Label(
            main_frame,
            text="اطلب من الطلاب مسح رمز QR التالي لتسجيل الحضور",
            style='Light.TLabel'
        )
        desc_label.pack(pady=(0, 20))
        
        # عرض صورة QR
        try:
            from PIL import Image, ImageTk
            
            qr_image = Image.open(qr_path)
            qr_image = qr_image.resize((350, 350), Image.LANCZOS)
            qr_photo = ImageTk.PhotoImage(qr_image)
            
            qr_label = tk.Label(main_frame, image=qr_photo, bg=ThemeManager.BG_LIGHT)
            qr_label.image = qr_photo  # حفظ مرجع للصورة
            qr_label.pack(pady=10)
        except Exception as e:
            # في حالة فشل عرض الصورة
            qr_error_label = ttk.Label(
                main_frame,
                text=f"تعذر عرض رمز QR: {str(e)}",
                style='Light.TLabel'
            )
            qr_error_label.pack(pady=10)
            
            # فتح الصورة في المتصفح
            webbrowser.open(f"file://{os.path.abspath(qr_path)}")
        
        # معلومات المدة
        time_frame = ttk.Frame(main_frame, style='Light.TFrame')
        time_frame.pack(pady=20, fill=tk.X)
        
        time_label = ttk.Label(
            time_frame,
            text=f"مدة التحضير: {duration_minutes} دقيقة",
            font=("Segoe UI", 12, "bold"),
            foreground=ThemeManager.TEXT_ACCENT,
            background=ThemeManager.BG_LIGHT
        )
        time_label.pack()
        
        # متغير لعرض الوقت المتبقي
        self.remaining_time_var = tk.StringVar(value="جاري بدء التحضير...")
        
        remaining_label = ttk.Label(
            time_frame,
            textvariable=self.remaining_time_var,
            style='Light.TLabel'
        )
        remaining_label.pack(pady=(5, 0))
        
        # زر إغلاق
        close_button = ttk.Button(
            main_frame,
            text="إنهاء التحضير",
            style='Primary.TButton',
            command=self.end_attendance_session
        )
        close_button.pack(pady=(20, 0), fill=tk.X)
        
        # منع إغلاق النافذة بالزر X
        self.qr_window.protocol("WM_DELETE_WINDOW", self.end_attendance_session)
    
    def attendance_timer_thread(self, duration_seconds):
        """مؤقت الحضور في خيط منفصل"""
        start_time = time.time()
        end_time = start_time + duration_seconds
        
        try:
            while time.time() < end_time:
                # حساب الوقت المتبقي
                remaining = end_time - time.time()
                minutes = int(remaining // 60)
                seconds = int(remaining % 60)
                
                # تحديث النص في واجهة المستخدم
                if self.qr_window and self.qr_window.winfo_exists():
                    self.remaining_time_var.set(f"الوقت المتبقي: {minutes:02d}:{seconds:02d}")
                
                # تحديث كل ثانية
                time.sleep(1)
            
            # انتهاء الوقت
            self.complete_attendance_session()
            
        except Exception as e:
            print(f"خطأ في مؤقت الحضور: {e}")
    
    def mark_student_attendance(self, student_name, student_id, validate_only=False):
        """تسجيل حضور الطالب"""
        if not self.excel_handler or not self.lecture_date:
            return False, "خطأ في النظام: لم يتم تهيئة نظام التحضير بشكل صحيح"
        
        return self.excel_handler.mark_attendance(student_name, student_id, self.lecture_date)
    
    def end_attendance_session(self):
        """إنهاء جلسة التحضير يدوياً"""
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من رغبتك في إنهاء جلسة التحضير؟"):
            self.complete_attendance_session()
    
    def complete_attendance_session(self):
        """إكمال جلسة التحضير وحفظ النتائج"""
        try:
            if self.excel_handler and self.lecture_date:
                # تسجيل الطلاب الغائبين
                absent_count = self.excel_handler.mark_all_absent(self.lecture_date)
                
                # حفظ الملف
                success, message = self.excel_handler.save_file()
                
                # عرض رسالة الإكمال
                self.show_completion_message(absent_count)
                
                # حذف ملفات الإخراج
                self.cleanup_files()
            
            # تنظيف الموارد
            self.cleanup_resources()
            
            # تحديث الحالة
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
        """عرض رسالة إكمال التحضير بتصميم محسن"""
        completion_dialog = tk.Toplevel(self.root)
        completion_dialog.title("اكتمال التحضير")
        completion_dialog.configure(bg=ThemeManager.BG_LIGHT)
        completion_dialog.resizable(False, False)
        
        # جعل النافذة مركزية
        self.center_window_on_parent(completion_dialog, 450, 300)
        
        # جعل النافذة مودال
        completion_dialog.transient(self.root)
        completion_dialog.grab_set()
        
        # إطار رئيسي
        main_frame = ttk.Frame(completion_dialog, style='Light.TFrame')
        main_frame.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
        
        # أيقونة النجاح
        success_label = ttk.Label(
            main_frame,
            text="✓",
            font=("Segoe UI", 48),
            foreground=ThemeManager.SUCCESS,
            background=ThemeManager.BG_LIGHT
        )
        success_label.pack(pady=(0, 10))
        
        # عنوان
        title_label = ttk.Label(
            main_frame,
            text="تم اكتمال التحضير بنجاح",
            font=("Segoe UI", 16, "bold"),
            foreground=ThemeManager.PRIMARY,
            background=ThemeManager.BG_LIGHT
        )
        title_label.pack(pady=(0, 20))
        
        # معلومات
        info_frame = ttk.Frame(main_frame, style='Light.TFrame')
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = f"تم تسجيل {absent_count} طالب كغائب"
        info_label = ttk.Label(
            info_frame,
            text=info_text,
            style='Light.TLabel'
        )
        info_label.pack()
        
        save_label = ttk.Label(
            info_frame,
            text="تم حفظ بيانات التحضير في ملف الإكسل",
            style='Light.TLabel'
        )
        save_label.pack(pady=(5, 0))
        
        # زر إغلاق
        close_button = ttk.Button(
            main_frame,
            text="إغلاق",
            style='Primary.TButton',
            command=completion_dialog.destroy
        )
        close_button.pack(fill=tk.X, pady=(10, 0))
        
        # تعيين زر الإغلاق كزر افتراضي
        completion_dialog.bind('<Return>', lambda event: completion_dialog.destroy())
    
    def cleanup_resources(self):
        """تنظيف الموارد المستخدمة"""
        # إغلاق نافذة QR
        if self.qr_window and self.qr_window.winfo_exists():
            self.qr_window.destroy()
            self.qr_window = None
        
        # قطع اتصال ngrok
        if self.ngrok_tunnel:
            try:
                ngrok.disconnect(self.ngrok_tunnel.public_url)
            except:
                pass
            self.ngrok_tunnel = None
        
        # إعادة تعيين المتغيرات
        self.excel_handler = None
        self.lecture_date = None
        self.timer_thread = None
    
    def cleanup_files(self):
        """حذف ملفات الإخراج"""
        try:
            output_files = [os.path.abspath(OUTPUT_FILENAME), os.path.abspath(QR_OUTPUT_FILENAME)]
            for file_path in output_files:
                if os.path.exists(file_path):
                    os.remove(file_path)
        except Exception as e:
            print(f"خطأ أثناء حذف الملفات: {e}")
    
    def show_styled_message(self, title, message, message_type="info"):
        """عرض رسالة بتصميم محسن"""
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
            text=f"MyTVTC - نظام التحضير الإلكتروني المتكامل\nالإصدار 2.0.0",
            style='Light.TLabel',
            justify=tk.CENTER
        )
        system_info.pack()
        
        # إضافة إطار منفصل للزر في أسفل النافذة
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