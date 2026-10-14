from datetime import datetime
import json
import re
import openpyxl
import requests

# Konfigurasi URL dan Path File
GAS_URL = "https://script.google.com/macros/s/AKfycbzWShHmhG0TNsnARV9cMbpHL2ludYiD1YTf999k0ZWBhCCSuK9P960TyFKlRFi2isO5VA/exec?action=getData"
JSON_FILE = "/tmp/gdrive_data.json"

def main():
    print("Mengambil data dari Google Spreadsheet...")
    
    # Melakukan HTTP GET request menggunakan pustaka requests
    response = requests.get(GAS_URL)
    print(f"HTTP Status Code: {response.status_code}")
    
    if response.status_code != 200 or not response.text.strip():
        print("ERROR: Gagal mengunduh data atau file kosong!")
        return
    
    # Simpan respon JSON ke file
    data = response.json()
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        
    print("--- 5 Baris Pertama Isi /tmp/gdrive_data.json ---")
    # Tampilkan cuplikan data JSON
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        lines = [f.readline() for _ in range(5)]
        print("".join(lines))
    print("------------------------------------------------")

    print("Memasukkan data ke file Excel lokal via Python...")

    # Kamus terjemahan manual untuk Hari dan Bulan ke Bahasa Indonesia
    DAYS_ID = {
        'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu', 
        'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
    }

    MONTHS_ID = {
        'January': 'Januari', 'February': 'Februari', 'March': 'Maret', 
        'April': 'April', 'May': 'Mei', 'June': 'Juni', 
        'July': 'Juli', 'August': 'Agustus', 'September': 'September', 
        'October': 'Oktober', 'November': 'November', 'December': 'Desember'
    }

    def format_date_if_matched(val):
        """Mendeteksi apakah nilai berupa tanggal dan mengubahnya ke Dddd, dd Mmmm yyyy"""
        if not isinstance(val, str):
            return val
        
        # Deteksi pola tanggal umum (YYYY-MM-DD atau ISO format dari GAS)
        date_pattern = r'^\d{4}-\d{2}-\d{2}'
        if re.match(date_pattern, val):
            try:
                dt = datetime.strptime(val[:10], '%Y-%m-%d')
                eng_day = dt.strftime('%A')
                eng_month = dt.strftime('%B')
                
                day_id = DAYS_ID.get(eng_day, eng_day)
                month_id = MONTHS_ID.get(eng_month, eng_month)
                
                return f"{day_id}, {dt.day:02d} {month_id} {dt.year}"
            except ValueError:
                return val
                
        return val

    def update_excel(file_path, sheet_mapping):
        try:
            wb = openpyxl.load_workbook(file_path)
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            
        sheet_name_map = {name.strip().lower(): name for name in wb.sheetnames}
            
        for json_key, target_sheet_name in sheet_mapping.items():
            rows = data.get(json_key, [])
            normalized_target = target_sheet_name.strip().lower()
            
            if normalized_target in sheet_name_map:
                ws = wb[sheet_name_map[normalized_target]]
            else:
                ws = wb.create_sheet(title=target_sheet_name)
                sheet_name_map[normalized_target] = target_sheet_name
            
            if ws.max_row >= 2:
                ws.delete_rows(2, ws.max_row - 1)
                
            for r_idx, row_data in enumerate(rows, start=2):
                for c_idx, val in enumerate(row_data, start=1):
                    # Deteksi dan ubah format tanggal jika cocok
                    formatted_val = format_date_if_matched(val)
                    
                    # KHUSUS SHEET "Anggota" KOLOM F (c_idx = 6): Paksa ubah ke string/text
                    if target_sheet_name.lower() == "anggota" and c_idx == 6:
                        if val is not None:
                            formatted_val = str(val)
                    
                    cell = ws.cell(row=r_idx, column=c_idx, value=formatted_val)
                    
                    # Pastikan tipe data cell diset sebagai string ('s') untuk kolom F di sheet Anggota
                    if target_sheet_name.lower() == "anggota" and c_idx == 6:
                        cell.data_type = 's'
                    
        wb.save(file_path)
        print(f"Berhasil memperbarui: {file_path}")

    # Proses pembaruan file Excel 1
    update_excel("aset/ST/data.xlsx", {
        "shST": "ST",
        "shSTTim": "Tim",
        "shSTAnggota": "Anggota"
    })

    # Proses pembaruan file Excel 2
    update_excel("aset/Lap/data.xlsx", {
        "shLap": "ST",
        "shLapTim": "Tim",
        "shLapAnggota": "Anggota"
    })

    print("Sinkronisasi data selesai!")

if __name__ == "__main__":
    main()
