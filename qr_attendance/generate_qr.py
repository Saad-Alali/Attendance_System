import qrcode
import json
import uuid
from datetime import datetime
import numpy as np

def generate_lecture_qr(lecture_name, lecture_date, web_url=None, output_path=None):
    session_code = str(uuid.uuid4())[:8]
    
    if isinstance(lecture_name, (np.integer, np.floating)):
        lecture_name = lecture_name.item()
    else:
        lecture_name = str(lecture_name)
    
    if web_url:
        # Create a QR code with the web URL directly (no JSON)
        qr_content = web_url
        print(f"Generating QR code with URL: {web_url}")
    else:
        # Create a QR code with JSON data (legacy mode)
        qr_data = {
            "lecture_name": lecture_name,
            "date": lecture_date.strftime("%Y-%m-%d"),
            "session_code": session_code,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        qr_content = json.dumps(qr_data)
        print("Generating QR code with JSON data (legacy mode)")
    
    # Create QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,  # Higher error correction
        box_size=10,
        border=4,
    )
    
    qr.add_data(qr_content)
    qr.make(fit=True)
    
    # Create image from QR code
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Save the image if output path is provided
    if output_path:
        img.save(output_path)
        print(f"QR code saved to: {output_path}")
    
    return img, session_code, output_path