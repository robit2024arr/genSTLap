import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

# Tentukan direktori download lokal
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
else:
    print("Membersihkan folder downloads...")
    for filename in os.listdir(download_dir):
        file_path = os.path.join(download_dir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"Gagal menghapus {file_path}: {e}")

# Konfigurasi Chrome Options
options = Options()

# Mode headless Wajib diaktifkan di Cloud Run karena container tidak punya display fisik
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--window-size=1920,1080")

# PENTING: Tentukan path binary Chromium bawaan apt-get
options.binary_location = "/usr/bin/chromium"

# Konfigurasi agar file terdownload otomatis ke folder 'downloads'
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

# Inisialisasi Driver menggunakan ChromeDriver sistem
service = Service("/usr/bin/chromedriver")
driver = webdriver.Chrome(service=service, options=options)

try:
    # ==========================================
    # 1. AKSES SITUS SIPONGI & DOWNLOAD TXT
    # ==========================================
    print("Memulai: Bot Sipongi...")
    print("Mengakses situs Sipongi...")
    driver.get("https://sipongi.gakkum.kehutanan.go.id/sebaran-titik-panas")

    time.sleep(4)
    wait = WebDriverWait(driver, 20)

    print("Memilih wilayah Kalimantan Tengah...")
    select_element = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "/html/body/div/div[1]/main/section[2]/div/div[2]/div[1]/div/div[2]/select",
            )
        )
    )

    wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "/html/body/div/div[1]/main/section[2]/div/div[2]/div[1]/div/div[2]/select/option[@value='13']",
            )
        )
    )

    dropdown = Select(select_element)
    dropdown.select_by_value("13")  # 13 = Kalimantan Tengah

    print("Mengklik tombol menu dropdown download...")
    dropdown_toggle = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div[1]/main/section[2]/div/div[2]/div[1]/div/div[5]/div/button",
            )
        )
    )
    driver.execute_script("arguments[0].click();", dropdown_toggle)

    print("Mengklik opsi download TXT...")
    download_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div[1]/main/section[2]/div/div[2]/div[1]/div/div[5]/div/ul/li[2]/a",
            )
        )
    )
    driver.execute_script("arguments[0].click();", download_btn)

    print("Menunggu proses download file .txt selesai...")
    downloaded_file_path = None
    timeout = 30
    start_time = time.time()

    while time.time() - start_time < timeout:
        files = os.listdir(download_dir)
        txt_files = [f for f in files if f.endswith(".txt")]
        if txt_files:
            txt_files.sort(
                key=lambda x: os.path.getmtime(os.path.join(download_dir, x)),
                reverse=True,
            )
            downloaded_file_path = os.path.join(download_dir, txt_files[0])
            break
        time.sleep(1)

    if not downloaded_file_path:
        raise Exception("Gagal mendownload file .txt dalam batas waktu yang ditentukan.")

    print(f"File berhasil diunduh di: {downloaded_file_path}")

    # ==========================================
    # 2. UPLOAD FILE KE GOOGLE APPS SCRIPT
    # ==========================================
    print("Mengakses Google Apps Script...")
    gas_url = "https://script.google.com/macros/s/AKfycbxQBLOGJHsUI5FJTLhrUXdhZf2c3ep8XjRBSuPcqOboI53BwpEWks0qH_gCTwO-a7FiLQ/exec"
    driver.get(gas_url)

    print("Menunggu halaman Google Apps Script termuat...")
    time.sleep(5)

    # CEK APAKAH SERVER GOOGLE MENGEMBALIKAN 'Service Unavailable'
    page_source = driver.page_source
    if "Service Unavailable" in page_source or "503" in page_source:
        raise Exception("[ERROR KONEKSI] Server Google Apps Script merespons: Service Unavailable. Silakan coba beberapa saat lagi.")

    def switch_to_element_context(driver, by, value):
        if driver.find_elements(by, value):
            return True

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for index, iframe in enumerate(iframes):
            try:
                driver.switch_to.frame(iframe)
                if switch_to_element_context(driver, by, value):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.parent_frame()
                continue
        return False

    print("Mencari elemen input file (id='fileInput') secara mendalam...")
    driver.switch_to.default_content()

    found_input = False
    start_search = time.time()
    while time.time() - start_search < 20:
        driver.switch_to.default_content()
        if switch_to_element_context(driver, By.ID, "fileInput"):
            found_input = True
            break
        time.sleep(2)

    if not found_input:
        raise Exception("Elemen dengan id='fileInput' tidak ditemukan di iframe manapun! Kemungkinan halaman Google Apps Script gagal dimuat atau menampilkan error.")

    file_input = driver.find_element(By.ID, "fileInput")
    driver.execute_script(
        "arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';",
        file_input,
    )

    file_input.send_keys(downloaded_file_path)
    print("Memasukkan file txt ke form input berhasil.")

    print("Mengklik tombol submit...")
    submit_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div/button"))
    )
    driver.execute_script("arguments[0].click();", submit_btn)

    # ==========================================
    # 3. MEMBACA STATUS & MENAMPILKAN KE TERMINAL
    # ==========================================
    print("Menunggu proses server Google Apps Script selesai...")
    status_text = ""
    timeout_status = 60
    start_status_time = time.time()

    while time.time() - start_status_time < timeout_status:
        try:
            status_div = driver.find_element(By.ID, "status")
            status_text = status_div.text
            if "Berhasil" in status_text or "Gagal" in status_text:
                break
        except Exception:
            pass
        time.sleep(2)

    print("\n================== HASIL STATUS ==================")
    print(status_text if status_text else "Waktu tunggu habis membaca status.")
    print("==================================================\n")

except Exception as e:
    print(f"Terjadi kesalahan: {e}")

finally:
    driver.quit()
    print("Browser ditutup.")
