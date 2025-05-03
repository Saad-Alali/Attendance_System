import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from project.utils import wait_for_download

def automate_browser(download_dir):
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    driver = webdriver.Chrome(options=chrome_options)
    
    driver.get("https://tvtc.gov.sa/ar/Departments/tvtcdepartments/Rayat/pages/E-Services.aspx")
    
    print("الرجاء تسجيل الدخول يدويًا...")
    
    target_tab = None
    main_window = driver.current_window_handle
    
    while True:
        time.sleep(2)
        current_handles = driver.window_handles
        
        for handle in current_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                current_url = driver.current_url
                if "rytfac.tvtc.gov.sa/FacultySelfService/ssb/GradeEntry" in current_url and "#/gradebook" in current_url:
                    target_tab = handle
                    break
        
        if target_tab:
            driver.switch_to.window(target_tab)
            print("تم العثور على صفحة الدرجات!")
            break
            
    print("تم الوصول إلى صفحة الدرجات. بدء العملية...")
    
    process_subjects(driver, download_dir)
    
    driver.quit()

def process_subjects(driver, download_dir):
    wait = WebDriverWait(driver, 20)
    
    try:
        time.sleep(2)
        print("البحث عن المواد...")
        
        subject_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[@data-name='subject']")))
        num_subjects = len(subject_rows)
        print(f"تم العثور على {num_subjects} مادة.")
        
        for i in range(num_subjects):
            print(f"معالجة المادة {i+1} من {num_subjects}...")
            subject_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[@data-name='subject']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", subject_rows[i])
            time.sleep(2)
            subject_rows[i].click()
            
            time.sleep(2)
            
            components_button = wait.until(EC.element_to_be_clickable((By.ID, "componentsButton")))
            components_button.click()
            
            time.sleep(2)
            
            print(f"بدء معالجة مكونات المادة {i+1}...")
            process_components(driver, wait, download_dir)
            
            print(f"العودة إلى قائمة المواد...")
            gradebook_tab = wait.until(EC.element_to_be_clickable((By.ID, "gradebook-tab")))
            gradebook_tab.click()
            
            time.sleep(2)
    except Exception as e:
        print(f"حدث خطأ أثناء معالجة المواد: {str(e)}")

def process_components(driver, wait, download_dir):
    try:
        time.sleep(5)  # Aumentamos el tiempo de espera para asegurar que la página cargue completamente
        print("البحث عن المكونات...")
        
        # Intenta obtener información sobre la página actual para depuración
        print(f"URL actual: {driver.current_url}")
        
        # Intentamos diferentes selectores para encontrar los componentes
        try:
            # Primer intento con selector más específico
            component_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='dataTable']//tr")))
            print(f"Encontrados {len(component_rows)} componentes con selector específico.")
        except Exception as e:
            print(f"No se encontraron componentes con selector específico: {str(e)}")
            try:
                # Segundo intento con selector más general
                component_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[td]")))
                print(f"Encontrados {len(component_rows)} componentes con selector general.")
            except Exception as e:
                print(f"No se encontraron componentes con selector general: {str(e)}")
                # Último intento con un selector muy general
                component_rows = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
                print(f"Encontrados {len(component_rows)} filas 'tr' en total.")
        
        print(f"تم العثور على {len(component_rows)} مكون.")
        
        # Si no hay componentes, tomemos una captura de pantalla para debug
        if len(component_rows) == 0:
            screenshot_path = os.path.join(download_dir, "debug_screenshot.png")
            driver.save_screenshot(screenshot_path)
            print(f"Se guardó una captura de pantalla en {screenshot_path}")
            print("No se encontraron componentes. Revisando html de la página...")
            print(driver.page_source[:1000])  # Muestra los primeros 1000 caracteres del HTML
            return
        
        # Proceso modificado para identificar componentes
        for i, row in enumerate(component_rows):
            try:
                print(f"Examinando fila {i+1}...")
                # Intentamos hacer clic directamente en la fila
                driver.execute_script("arguments[0].scrollIntoView(true);", row)
                time.sleep(2)
                row.click()
                
                time.sleep(2)
                
                # Intentamos encontrar y hacer clic en el icono de herramientas
                try:
                    tools_icon = driver.find_element(By.ID, "tools")
                    if tools_icon.is_displayed():
                        print("Encontró icono de herramientas, procesando componente...")
                        process_component(driver, wait, download_dir)
                        
                        # Regresamos a la lista de componentes
                        try:
                            back_button = driver.find_element(By.XPATH, "//a[contains(@href, '#/components')]")
                            back_button.click()
                            time.sleep(2)
                        except Exception as e:
                            print(f"Error al regresar a la lista de componentes: {str(e)}")
                            
                        # Actualizamos la lista de filas
                        component_rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[td]")))
                    else:
                        print("Icono de herramientas no visible, pasando a la siguiente fila")
                except Exception as e:
                    print(f"No se encontró icono de herramientas: {str(e)}")
                    
            except Exception as e:
                print(f"Error al procesar fila {i+1}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"حدث خطأ أثناء معالجة المكونات: {str(e)}")
        # Captura de pantalla en caso de error
        screenshot_path = os.path.join(download_dir, "error_screenshot.png")
        driver.save_screenshot(screenshot_path)
        print(f"Se guardó una captura de pantalla del error en {screenshot_path}")

def process_component(driver, wait, download_dir):
    try:
        print("  البحث عن أيقونة الأدوات...")
        tools_icon = wait.until(EC.element_to_be_clickable((By.ID, "tools")))
        driver.execute_script("arguments[0].scrollIntoView(true);", tools_icon)
        time.sleep(2)
        tools_icon.click()
        
        print("  اختيار 'تصدير القالب'...")
        time.sleep(2)
        export_template = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='تصدير القالب']")))
        export_template.click()
        
        print("  اختيار 'جداول البيانات إكسل'...")
        time.sleep(2)
        excel_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='xlsxType' and contains(text(), 'جداول البيانات إكسل')]")))
        excel_option.click()
        
        print("  الضغط على 'تصدير'...")
        time.sleep(2)
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@ng-click, 'ok()') and contains(text(), 'تصدير')]")))
        export_button.click()
        
        print("  انتظار اكتمال التحميل...")
        time.sleep(2)
        wait_for_download(download_dir)
        print("  تم التحميل بنجاح.")
        time.sleep(2)
    except Exception as e:
        print(f"  حدث خطأ أثناء معالجة المكون: {str(e)}")