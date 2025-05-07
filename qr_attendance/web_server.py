from flask import Flask, request, render_template, jsonify, make_response
import threading
import socket
import os
import json
import re
from datetime import datetime
from security.fingerprint import DeviceFingerprint

app = Flask(__name__, template_folder=os.path.join(os.path.dirname(__file__), 'templates'))

attendance_data = {
    'session_code': None,
    'callback_function': None,
    'students': [],
    'fingerprint_failures': {}  # Para registrar fallos de verificación por estudiante
}

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return socket.gethostbyname(socket.gethostname())

def get_client_device_info(request):
    user_agent = request.user_agent.string
    device_info = {}
    
    if 'Mobile' in user_agent:
        device_info['device_type'] = 'هاتف محمول'
    elif 'Tablet' in user_agent:
        device_info['device_type'] = 'جهاز لوحي'
    else:
        device_info['device_type'] = 'كمبيوتر'
    
    if 'Chrome' in user_agent and 'Edg' not in user_agent:
        device_info['browser'] = 'كروم'
    elif 'Firefox' in user_agent:
        device_info['browser'] = 'فايرفوكس'
    elif 'Safari' in user_agent and 'Chrome' not in user_agent:
        device_info['browser'] = 'سفاري'
    elif 'Edg' in user_agent:
        device_info['browser'] = 'إيدج'
    elif 'MSIE' in user_agent or 'Trident' in user_agent:
        device_info['browser'] = 'انترنت اكسبلورر'
    else:
        device_info['browser'] = 'متصفح غير معروف'
    
    if 'Android' in user_agent:
        device_info['os'] = 'أندرويد'
        version_match = re.search(r'Android\s([0-9\.]+)', user_agent)
        if version_match:
            device_info['os_version'] = version_match.group(1)
    elif 'iPhone' in user_agent or 'iPad' in user_agent or 'iPod' in user_agent:
        device_info['os'] = 'iOS'
        version_match = re.search(r'OS\s([0-9_]+)', user_agent)
        if version_match:
            device_info['os_version'] = version_match.group(1).replace('_', '.')
    elif 'Windows' in user_agent:
        device_info['os'] = 'ويندوز'
        version_match = re.search(r'Windows NT\s([0-9\.]+)', user_agent)
        if version_match:
            nt_version = version_match.group(1)
            windows_versions = {
                '10.0': '10/11',
                '6.3': '8.1',
                '6.2': '8',
                '6.1': '7',
                '6.0': 'فيستا',
                '5.2': 'XP 64-bit',
                '5.1': 'XP',
            }
            device_info['os_version'] = windows_versions.get(nt_version, nt_version)
    elif 'Mac OS X' in user_agent:
        device_info['os'] = 'ماك'
        version_match = re.search(r'Mac OS X\s([0-9_\.]+)', user_agent)
        if version_match:
            device_info['os_version'] = version_match.group(1).replace('_', '.')
    elif 'Linux' in user_agent:
        device_info['os'] = 'لينكس'
    else:
        device_info['os'] = 'نظام غير معروف'
    
    device_info['ip_address'] = request.remote_addr
    device_info['user_agent'] = user_agent
    
    try:
        js_data = json.loads(request.form.get('client_device_data', '{}'))
        if js_data:
            if 'screenWidth' in js_data and 'screenHeight' in js_data:
                device_info['screen'] = f"{js_data['screenWidth']}×{js_data['screenHeight']}"
            if 'language' in js_data:
                device_info['language'] = js_data['language']
            if 'platform' in js_data:
                device_info['platform'] = js_data['platform']
            if 'webgl_vendor' in js_data:
                device_info['graphics'] = js_data['webgl_vendor']
            if 'timezone' in js_data:
                device_info['timezone'] = js_data['timezone']
    except:
        pass
    
    return device_info

@app.route('/')
def index():
    return "QR Attendance System Server Running"

@app.route('/attendance')
def attendance_form():
    session_code = request.args.get('session', '')
    if session_code != attendance_data['session_code']:
        return "Invalid or expired session"
    
    return render_template('attendance_form.html', session_code=session_code)

@app.route('/attendance_confirmed')
def attendance_confirmed():
    return render_template('attendance_confirmed.html')

@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    if request.method == 'POST':
        session_code = request.form.get('session_code')
        student_name = request.form.get('student_name')
        student_id = request.form.get('student_id')
        
        if not all([session_code, student_name, student_id]):
            response = make_response(jsonify({'status': 'error', 'message': 'يجب ملء جميع الحقول'}))
            response.status_code = 400  # Bad Request
            return response
        
        if session_code != attendance_data['session_code']:
            response = make_response(jsonify({'status': 'error', 'message': 'الجلسة غير صالحة أو منتهية الصلاحية'}))
            response.status_code = 403  # Forbidden
            return response
        
        # Verificar si el estudiante ya está registrado con un error de huella
        student_key = f"{student_name}_{student_id}"
        if student_key in attendance_data.get('fingerprint_failures', {}):
            error_info = attendance_data['fingerprint_failures'][student_key]
            response = make_response(jsonify({
                'status': 'error', 
                'message': f"تم رفض التسجيل مسبقاً بسبب: {error_info['message']}"
            }))
            response.status_code = 403  # Forbidden
            return response
        
        # Validar primero si el estudiante existe en el sistema
        validation_result = None
        if attendance_data['callback_function']:
            is_valid, message = attendance_data['callback_function'](student_name, student_id, validate_only=True)
            validation_result = (is_valid, message)
            
            if not is_valid:
                response = make_response(jsonify({'status': 'error', 'message': message}))
                response.status_code = 404  # Not Found
                return response
        
        # Verificar la huella digital del dispositivo
        fingerprint = DeviceFingerprint()
        is_fingerprint_valid, fingerprint_message = fingerprint.verify_student(student_name, request)
        
        # CAMBIO CRÍTICO: Si la huella digital no es válida, rechazar completamente
        if not is_fingerprint_valid:
            # Registrar el intento fallido para auditoría
            device_info = get_client_device_info(request)
            print(f"⚠️ Error de verificación de huella: {student_name} ({student_id}) - {fingerprint_message}")
            print(f"  Desde dispositivo: {device_info.get('device_type', 'desconocido')} - {device_info.get('os', 'desconocido')} - IP: {device_info.get('ip_address', 'desconocido')}")
            
            # Registrar este estudiante como fallido para evitar intentos repetidos
            attendance_data.setdefault('fingerprint_failures', {})[student_key] = {
                'message': fingerprint_message,
                'timestamp': datetime.now().isoformat(),
                'device_info': device_info
            }
            
            # Devolver error HTTP 403 Forbidden para que el cliente lo maneje correctamente
            response = make_response(jsonify({'status': 'error', 'message': fingerprint_message}))
            response.status_code = 403  # Forbidden
            return response
        
        # Si llegamos aquí, la huella es válida, recopilamos información del dispositivo
        device_info = get_client_device_info(request)
        
        student_data = {
            'name': student_name,
            'id': student_id,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'device_info': {
                'نوع الجهاز': device_info.get('device_type', 'غير معروف'),
                'نظام التشغيل': f"{device_info.get('os', 'غير معروف')} {device_info.get('os_version', '')}",
                'المتصفح': device_info.get('browser', 'غير معروف'),
                'دقة الشاشة': device_info.get('screen', 'غير معروف'),
                'اللغة': device_info.get('language', 'غير معروف'),
                'المنطقة الزمنية': device_info.get('timezone', 'غير معروف'),
                'عنوان IP': device_info.get('ip_address', 'غير معروف')
            }
        }
        
        print(f"معلومات جهاز الطالب {student_name} ({student_id}):")
        for key, value in student_data['device_info'].items():
            print(f"  {key}: {value}")
        
        # Registrar los datos del estudiante
        attendance_data['students'].append(student_data)
        
        # Marcar la asistencia en Excel
        if attendance_data['callback_function'] and validation_result and validation_result[0]:
            result = attendance_data['callback_function'](student_name, student_id)
            if result and isinstance(result, tuple) and len(result) >= 2 and not result[0]:
                # Si la función de callback devuelve un error, regresar error al cliente
                response = make_response(jsonify({'status': 'error', 'message': result[1]}))
                response.status_code = 400
                return response
                
            print(f"✅ Asistencia registrada para: {student_name} ({student_id})")
        
        return jsonify({'status': 'success', 'message': 'تم تسجيل حضورك بنجاح'})

def start_server(host=None, port=5000, session_code=None, callback=None):
    if host is None:
        host = get_local_ip()
    
    attendance_data['session_code'] = session_code
    attendance_data['callback_function'] = callback
    attendance_data['students'] = []
    attendance_data['fingerprint_failures'] = {}
    
    server_url = f"http://{host}:{port}"
    attendance_url = f"{server_url}/attendance?session={session_code}"
    
    threading.Thread(target=lambda: app.run(host=host, port=port, debug=False), daemon=True).start()
    
    return server_url, attendance_url

def stop_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Not running with the Werkzeug Server')
    func()
    return "Server shutting down..."