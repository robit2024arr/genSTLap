# 1. Buat virtual environment bernama 'venv' di folder kerja Anda
python3 -m venv venv

# 2. Aktifkan virtual environment
source venv/bin/activate

# 3. Install pustaka yang dibutuhkan di dalam venv (sekarang pip bisa digunakan dengan bebas)
pip install pandas python-docx openpyxl prompt-toolkit 
pip install prompt-toolkit --break-system-packages
