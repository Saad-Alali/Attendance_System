INITIAL_URL = "https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx"
TARGET_URL = "https://rytfac.tvtc.gov.sa/FacultySelfService/ssb/facultyAttendanceTracking#!/markAttendance"

TOOLS_BUTTON_ID = "tools"
EXPORT_TOOL_ITEM_ID = "export-tool-item"

EXPORT_FILE_TYPE = "Excel (.xlsx)"
EXPORT_RANGE = "جميع الطلاب"
EXPORT_DATES = "كل أيام المحاضرات"

import datetime

# تعريف اسم ملف الإخراج مع إضافة التاريخ
def get_output_filename():
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return f"Attendance_{today}.xlsx"

OUTPUT_FILENAME = get_output_filename()


OUTPUT_PATH = "~/Desktop/"