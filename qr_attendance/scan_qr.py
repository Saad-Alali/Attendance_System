import cv2
import json
from datetime import datetime
from pyzbar.pyzbar import decode
import numpy as np

class QRScanner:
    def __init__(self):
        self.cap = None
    
    def start_camera(self):
        try:
            self.cap = cv2.VideoCapture(0)
            return self.cap.isOpened()
        except Exception as e:
            print(f"Camera error: {e}")
            return False
    
    def scan_qr_code(self, timeout=60):
        if not self.cap or not self.cap.isOpened():
            if not self.start_camera():
                raise Exception("Failed to open camera")
        
        start_time = datetime.now()
        
        print("Scanning for QR code... Press 'q' to exit")
        
        while (datetime.now() - start_time).total_seconds() < timeout:
            ret, frame = self.cap.read()
            
            if not ret:
                continue
            
            cv2.imshow('QR Scanner', frame)
            
            decoded_objects = decode(frame)
            
            for obj in decoded_objects:
                try:
                    data_json = obj.data.decode('utf-8')
                    qr_data = json.loads(data_json)
                    
                    print("\nQR code scanned successfully!")
                    
                    return qr_data
                except Exception as e:
                    print(f"Error reading QR code: {e}")
            
            key = cv2.waitKey(1) & 0xFF
            if key == ord('q'):
                break
        
        return None
    
    def release_resources(self):
        if self.cap and self.cap.isOpened():
            self.cap.release()
        cv2.destroyAllWindows()
    
    def get_student_info(self):
        print("\n=== Attendance Registration ===")
        student_name = input("Full Name: ")
        
        while True:
            try:
                student_id = input("Student ID: ")
                if student_id.strip():
                    break
                print("Please enter a valid student ID")
            except Exception:
                print("Please enter a valid student ID")
        
        return student_name, student_id
    
    def verify_session(self, scanned_session_code, original_session_code):
        return scanned_session_code == original_session_code