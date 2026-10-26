FROM python:3.10-slim

WORKDIR /app

# Install system dependencies, Chromium, dan Chromium Driver (tanpa libgconf-2-4)
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    wget \
    curl \
    unzip \
    chromium \
    chromium-driver \
    fonts-liberation \
    libnss3 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# Salin requirements.txt dan install dependencies Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Salin seluruh file proyek ke dalam container
COPY . .

# Berikan port default untuk Cloud Run (port 8080)
ENV PORT=8080

# Jalankan aplikasi menggunakan uvicorn
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
