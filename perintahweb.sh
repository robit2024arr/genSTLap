#!/bin/bash

# ==============================================================================
# Script Peluncur Web App Generator Laporan & Surat Tugas (BAPPERIDA)
# Menggantikan perintah.sh dengan antarmuka Web & REST API
# ==============================================================================

# Pastikan script berjalan dari direktori tempat script ini berada
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "===================================================="
echo "    MEMULAI WEB APP GENERATOR ST & LAPORAN         "
echo "===================================================="

# Cek virtual environment lokal
if [ -f "./venv/bin/uvicorn" ]; then
    PYTHON_CMD="./venv/bin/python"
    UVICORN_CMD="./venv/bin/uvicorn"
elif command -v uvicorn >/dev/null 2>&1; then
    PYTHON_CMD="python3"
    UVICORN_CMD="uvicorn"
else
    echo "[INFO] Menyiapkan environment..."
    python3 -m venv --system-site-packages venv
    ./venv/bin/pip install -r requirements.txt
    PYTHON_CMD="./venv/bin/python"
    UVICORN_CMD="./venv/bin/uvicorn"
fi

PORT=8000
echo ""
echo "🚀 Server Web App berhasil diinisialisasi!"
echo "----------------------------------------------------"
echo "👉 Buka di Browser Anda   : http://localhost:$PORT"
echo "👉 Dokumentasi REST API    : http://localhost:$PORT/docs"
echo "----------------------------------------------------"
echo "Tekan [Ctrl + C] kapan saja untuk menghentikan server."
echo "===================================================="
echo ""

# Jalankan Uvicorn Server (tanpa --reload agar server tidak restart saat file excel/docx dibuat)
$UVICORN_CMD app:app --host 0.0.0.0 --port $PORT --ws none
