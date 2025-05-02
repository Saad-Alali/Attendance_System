COLUMN_NAMES = {
    "date": "أيام المحاضرات",
    "full_name": "اسم الطالب",
    "student_id": "الرقم الجامعي للطالب",
    "attendance": "مؤشر الحضور",
    "expected_hours": "الساعات المتوقعة",
    "actual_hours": "الساعات الفعلية",
    "absence_hours": "ساعات الغياب",
    "authorized_absence": "غياب مصرح به",
    
    "lecture_name": "رمز الفصل الدراسي",
    "section": "الرقم المرجعي للمقرر"
}

ATTENDANCE_STATUS = {
    "present": "حاضر",
    "absent": "غائب"
}

AUTHORIZED_ABSENCE = {
    "no": "لا",
    "yes": "نعم"
}

import datetime

def get_qr_output_filename():
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return f"QR_Code_{today}.png"

QR_OUTPUT_FILENAME = get_qr_output_filename()

def get_output_filename():
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return f"Attendance_{today}.xlsx"

OUTPUT_FILENAME = get_output_filename()

QR_OUTPUT_FILENAME = "lecture_qr.png"

ATTENDANCE_DURATION = 15 * 60