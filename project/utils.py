import time
import os

def wait_for_download(download_dir, timeout=60):
    start_time = time.time()
    
    while True:
        if time.time() - start_time > timeout:
            return False
        
        files = [f for f in os.listdir(download_dir) if f.endswith('.crdownload') or f.endswith('.tmp')]
        
        if not files:
            time.sleep(1)
            return True
        
        time.sleep(1)