import pandas as pd

def validate_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        
        required_columns = [
            'رمز الفصل الدراسي', 
            'الرقم المرجعي للمقرر', 
            'الاسم الكامل', 
            'الرقم الجامعي للطالب', 
            'التخصص', 
            'LEC2', 
            'LEC1', 
            'LAB2', 
            'LAB1'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return False, f"الأعمدة التالية مفقودة: {', '.join(missing_columns)}"
        
        return True, "تم التحقق من الملف بنجاح"
    except Exception as e:
        return False, str(e)