import json
import os
import hashlib
from datetime import datetime

class DeviceFingerprint:
    def __init__(self, json_file_path='security/device_fingerprints.json'):
        self.json_file_path = json_file_path
        
        os.makedirs(os.path.dirname(json_file_path), exist_ok=True)
        
        if not os.path.exists(json_file_path):
            self._initialize_json_file()
        else:
            self._validate_json_file()
    
    def _initialize_json_file(self):
        with open(self.json_file_path, 'w', encoding='utf-8') as f:
            json.dump({"devices": {}, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}, f, ensure_ascii=False, indent=4)
        print(f"JSON fingerprint file initialized at {self.json_file_path}")
    
    def _validate_json_file(self):
        try:
            with open(self.json_file_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if not content:
                    self._initialize_json_file()
                else:
                    try:
                        data = json.loads(content)
                        if not isinstance(data, dict) or "devices" not in data:
                            updated_data = {
                                "devices": data if isinstance(data, dict) else {}, 
                                "device_mappings": {},
                                "metadata": {"updated": datetime.now().isoformat()}
                            }
                            with open(self.json_file_path, 'w', encoding='utf-8') as f:
                                json.dump(updated_data, f, ensure_ascii=False, indent=4)
                            print("Updated fingerprint file to new format")
                        elif "device_mappings" not in data:
                            data["device_mappings"] = {}
                            data["metadata"]["updated"] = datetime.now().isoformat()
                            with open(self.json_file_path, 'w', encoding='utf-8') as f:
                                json.dump(data, f, ensure_ascii=False, indent=4)
                            print("Added device_mappings to fingerprint file")
                    except json.JSONDecodeError:
                        backup_path = f"{self.json_file_path}.bak"
                        try:
                            import shutil
                            shutil.copy2(self.json_file_path, backup_path)
                            print(f"Backed up invalid JSON file to {backup_path}")
                        except Exception:
                            pass
                        self._initialize_json_file()
        except Exception as e:
            print(f"Error validating JSON file: {e}")
            self._initialize_json_file()
    
    def get_device_fingerprint(self, request=None):
        if request is None:
            dummy_info = {
                'timestamp': str(datetime.now().timestamp()),
                'random': str(os.urandom(8).hex())
            }
            primary_str = json.dumps(dummy_info, sort_keys=True)
            primary_fp = hashlib.sha256(primary_str.encode()).hexdigest()
            return {
                'primary': primary_fp,
                'secondary': primary_fp,
                'hardware': primary_fp,
                'raw': dummy_info
            }
        else:
            return self.get_browser_fingerprint(request)
    
    def get_browser_fingerprint(self, request):
        user_agent = request.user_agent.string
        ip_address = request.remote_addr
        
        try:
            device_data = json.loads(request.form.get('client_device_data', '{}'))
        except:
            device_data = {}
        
        primary_info = {
            'user_agent': user_agent,
            'platform': device_data.get('platform', ''),
            'webgl_renderer': device_data.get('webgl_renderer', ''),
        }
        
        secondary_info = {
            'screen_width': device_data.get('screenWidth', ''),
            'screen_height': device_data.get('screenHeight', ''),
            'color_depth': device_data.get('colorDepth', ''),
            'pixel_ratio': device_data.get('pixelRatio', ''),
            'language': device_data.get('language', ''),
            'timezone': device_data.get('timezone', ''),
            'ip_address': ip_address
        }
        
        hardware_info = {
            'platform': device_data.get('platform', ''),
            'webgl_renderer': device_data.get('webgl_renderer', ''),
            'webgl_vendor': device_data.get('webgl_vendor', ''),
            'screen_width': device_data.get('screenWidth', ''),
            'screen_height': device_data.get('screenHeight', ''),
            'color_depth': device_data.get('colorDepth', ''),
            'pixel_ratio': device_data.get('pixelRatio', ''),
        }
        
        primary_str = json.dumps(primary_info, sort_keys=True)
        primary_fp = hashlib.sha256(primary_str.encode()).hexdigest()
        
        all_info = {**primary_info, **secondary_info}
        all_str = json.dumps(all_info, sort_keys=True)
        secondary_fp = hashlib.sha256(all_str.encode()).hexdigest()
        
        hardware_str = json.dumps(hardware_info, sort_keys=True)
        hardware_fp = hashlib.sha256(hardware_str.encode()).hexdigest()
        
        return {
            'primary': primary_fp,
            'secondary': secondary_fp,
            'hardware': hardware_fp,
            'raw': all_info
        }
    
    def register_student(self, student_name, request=None):
        if request:
            fingerprints = self.get_browser_fingerprint(request)
        else:
            fingerprints = self.get_device_fingerprint()
        
        primary_fp = fingerprints['primary']
        secondary_fp = fingerprints['secondary']
        hardware_fp = fingerprints['hardware']
        
        data = self._read_data()
        devices = data.get('devices', {})
        device_mappings = data.get('device_mappings', {})
        
        if primary_fp in devices:
            registered_name = devices[primary_fp]['student']
            if registered_name == student_name:
                return True, f"تم التحقق: الطالب {student_name} مسجل بالفعل من هذا الجهاز"
            else:
                return False, f"خطأ: هذا الجهاز مسجل باسم {registered_name}. لا يمكن استخدام نفس الجهاز لطالب آخر"
        
        if hardware_fp in device_mappings:
            associated_student = device_mappings[hardware_fp]
            if associated_student != student_name:
                return False, f"خطأ: هذا الجهاز مستخدم بواسطة طالب آخر ({associated_student})"
        
        for fp, info in devices.items():
            if info.get('hardware') == hardware_fp and info.get('student') != student_name:
                return False, f"خطأ: هذا الجهاز (أو جهاز مشابه جداً) مسجل باسم {info.get('student')}."
            
            if info.get('secondary') == secondary_fp and info.get('student') != student_name:
                return False, f"خطأ: هذا الجهاز (أو جهاز مشابه جداً) مسجل باسم {info.get('student')}."
        
        devices[primary_fp] = {
            'student': student_name,
            'secondary': secondary_fp,
            'hardware': hardware_fp,
            'registered_at': datetime.now().isoformat(),
            'details': fingerprints['raw']
        }
        
        device_mappings[hardware_fp] = student_name
        
        data['devices'] = devices
        data['device_mappings'] = device_mappings
        
        if self._save_data(data):
            print(f"تم تسجيل الجهاز للطالب {student_name}")
            return True, f"تم تسجيل الطالب {student_name} بنجاح مع بصمة هذا الجهاز"
        else:
            return False, "خطأ في النظام: لا يمكن تسجيل الجهاز"
    
    def verify_student(self, student_name, request=None):
        if request:
            fingerprints = self.get_browser_fingerprint(request)
        else:
            fingerprints = self.get_device_fingerprint()
        
        primary_fp = fingerprints['primary']
        secondary_fp = fingerprints['secondary']
        hardware_fp = fingerprints['hardware']
        
        data = self._read_data()
        devices = data.get('devices', {})
        device_mappings = data.get('device_mappings', {})
        
        if primary_fp in devices:
            registered_name = devices[primary_fp]['student']
            if registered_name != student_name:
                return False, f"خطأ: هذا الجهاز مسجل باسم طالب آخر ({registered_name})"
            else:
                return True, f"تم التحقق بنجاح: الطالب {student_name} يستخدم الجهاز المسجل"
        
        if hardware_fp in device_mappings:
            associated_student = device_mappings[hardware_fp]
            if associated_student != student_name:
                return False, f"خطأ: هذا الجهاز مستخدم بواسطة طالب آخر ({associated_student})"
        
        for fp, info in devices.items():
            if info.get('hardware') == hardware_fp:
                if info.get('student') == student_name:
                    devices[primary_fp] = {
                        'student': student_name,
                        'secondary': secondary_fp,
                        'hardware': hardware_fp,
                        'registered_at': datetime.now().isoformat(),
                        'updated_at': datetime.now().isoformat(),
                        'details': fingerprints['raw']
                    }
                    self._save_data(data)
                    return True, f"تم التحقق: الطالب {student_name} يستخدم نفس الجهاز المسجل"
                else:
                    return False, f"خطأ: هذا الجهاز مسجل باسم {info.get('student')}"
            
            if info.get('secondary') == secondary_fp:
                if info.get('student') == student_name:
                    devices[primary_fp] = devices.pop(fp)
                    devices[primary_fp]['updated_at'] = datetime.now().isoformat()
                    devices[primary_fp]['details'] = fingerprints['raw']
                    self._save_data(data)
                    return True, f"تم التحقق: الطالب {student_name} يستخدم جهاز مشابه للمسجل"
                else:
                    return False, f"خطأ: جهاز مشابه مسجل باسم {info.get('student')}"
        
        return False, "هذا الجهاز غير مسجل بعد"
    
    def _read_data(self):
        try:
            with open(self.json_file_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if not content:
                    return {"devices": {}, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}
                try:
                    data = json.loads(content)
                    if not isinstance(data, dict):
                        return {"devices": {}, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}
                    if "devices" not in data:
                        return {"devices": data, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}
                    if "device_mappings" not in data:
                        data["device_mappings"] = {}
                    return data
                except json.JSONDecodeError:
                    print("محتوى JSON غير صالح، إعادة تعيين البيانات")
                    return {"devices": {}, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}
        except Exception as e:
            print(f"خطأ في قراءة ملف البصمات: {e}")
            return {"devices": {}, "device_mappings": {}, "metadata": {"created": datetime.now().isoformat()}}
    
    def _save_data(self, data):
        try:
            if "metadata" not in data:
                data["metadata"] = {}
            data["metadata"]["updated"] = datetime.now().isoformat()
            
            with open(self.json_file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"خطأ في حفظ البصمات: {e}")
            return False