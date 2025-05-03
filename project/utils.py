import time
import os

def wait_for_download(download_dir, timeout=30):
    start_time = time.time()
    
    while True:
        if time.time() - start_time > timeout:
            return False
        
        files = [f for f in os.listdir(download_dir) if f.endswith('.crdownload') or f.endswith('.tmp')]
        
        if not files:
            time.sleep(0.5)
            return True
        
        time.sleep(0.5)