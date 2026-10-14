#!/bin/bash

# Aktifkan mode debug untuk melacak eksekusi
set -x

GAS_URL="https://script.google.com/macros/s/AKfycbzWShHmhG0TNsnARV9cMbpHL2ludYiD1YTf999k0ZWBhCCSuK9P960TyFKlRFi2isO5VA/exec?action=getData"
JSON_FILE="/tmp/gdrive_data.json"
LOG_FILE="/tmp/gdrive_error.log"

echo "Mengambil data dari Google Spreadsheet..."

HTTP_CODE=$(curl -s -L -w "%{http_code}" "$GAS_URL" -o "$JSON_FILE")
echo "HTTP Status Code: $HTTP_CODE"

echo "--- 5 Baris Pertama Isi /tmp/gdrive_data.json ---"
head -n 5 "$JSON_FILE"
echo "------------------------------------------------"

if [ ! -s "$JSON_FILE" ]; then
    echo "ERROR: File JSON kosong atau gagal diunduh!"
    exit 1
fi

echo "Memasukkan data ke file Excel lokal via Python..."

python3 2> "$LOG_FILE" << 'EOF'
import json
import openpyxl
from datetime import datetime
import re

with open("/tmp/gdrive_data.json", "r") as f:
    data = json.load(f)

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
    # Contoh: 2026-09-16 atau 2026-09-16T00:00:00.000Z
    date_pattern = r'^\d{4}-\d{2}-\d{2}'
    if re.match(date_pattern, val):
        try:
            # Ambil bagian tanggal utamanya saja (10 karakter pertama)
            dt = datetime.strptime(val[:10], '%Y-%m-%d')
            
            # Format ke hari, tanggal, bulan, tahun bahasa inggris dulu untuk diterjemahkan
            eng_day = dt.strftime('%A')
            eng_month = dt.strftime('%B')
            
            day_id = DAYS_ID.get(eng_day, eng_day)
            month_id = MONTHS_ID.get(eng_month, eng_month)
            
            # Susun sesuai format Dddd, dd Mmmm yyyy (Contoh: Rabu, 16 September 2026)
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
                ws.cell(row=r_idx, column=c_idx, value=formatted_val)
                
    wb.save(file_path)
    print(f"Berhasil memperbarui: {file_path}")

update_excel("aset/ST/data.xlsx", {
    "shST": "ST",
    "shSTTim": "Tim",
    "shSTAnggota": "Anggota"
})

update_excel("aset/Lap/data.xlsx", {
    "shLap": "ST",
    "shLapTim": "Tim",
    "shLapAnggota": "Anggota"
})
EOF

PYTHON_EXIT_CODE=$?
if [ $PYTHON_EXIT_CODE -ne 0 ]; then
    echo "ERROR: Python gagal memproses data!"
    echo "--- LOG ERROR PYTHON ---"
    cat "$LOG_FILE"
    echo "------------------------"
    exit 1
fi

echo "Sinkronisasi data selesai!"
set +x
