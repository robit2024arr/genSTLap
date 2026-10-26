# Gunakan base image Python resmi
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# 1. Install dependensi sistem yang dibutuhkan oleh Chrome, Selenium, dan build tools (gcc)
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    wget \
    curl \
    gnupg \
    unzip \
    libxi6 \
    libgconf-2-4 \
    libnss3 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libasound2 \
    libpango-1.0-0 \
    libcairo2 \
    fonts-liberation \
    xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# 2. Install Google Chrome Stable dan ChromeDriver yang versinya sinkron secara otomatis
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg && \
    echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && apt-get install -y google-chrome-stable && \
    CHROME_VERSION=$(google-chrome --version | awk '{print $3}' | cut -d'.' -f1) && \
    echo "Installed Chrome version major: $CHROME_VERSION" && \
    # Download ChromeDriver versi yang sesuai dengan Chrome stable
    LATEST_DRIVER_VERSION=$(curl -s "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_$CHROME_VERSION") && \
    wget -q "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/$LATEST_DRIVER_VERSION/linux64/chromedriver-linux64.zip" && \
    unzip chromedriver-linux64.zip && \
    mv chromedriver-linux64/chromedriver /usr/bin/chromedriver && \
    chmod +x /usr/bin/chromedriver && \
    rm -rf chromedriver-linux64.zip chromedriver-linux64 && \
    rm -rf /var/lib/apt/lists/*

# Salin requirements.txt dan install dependencies Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Salin seluruh file proyek ke dalam container
COPY . .

# Berikan port default untuk Cloud Run (port 8080)
ENV PORT=8080

# Jalankan aplikasi menggunakan uvicorn
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
