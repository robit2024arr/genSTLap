FROM python:3.10-slim

WORKDIR /app

# Install system dependencies, Chromium, dan Chromium Driver langsung dari repository Debian
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    wget \
    curl \
    unzip \
    chromium \
    chromium-driver \
    fonts-liberation \
    libnss3 \
    libgconf-2-4 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# Salin requirements.txt dan install dependencies Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Salin seluruh file proyek ke dalam container
COPY . .

# Berikan port default untuk Cloud Run (port 8080)
ENV PORT=8080

# Jalankan aplikasi menggunakan uvicorn (sesuaikan 'app:app' jika nama file/instance Anda berbeda)
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
