# Dockerfile สำหรับ Hugging Face Spaces (Docker SDK)
# รัน Streamlit app บน port 7860 (HF Spaces convention)

FROM python:3.11-slim

# ตั้ง working directory
WORKDIR /app

# ติดตั้ง system dependencies ที่จำเป็น (น้อยที่สุด)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# ก๊อป requirements ก่อน เพื่อให้ Docker cache ทำงาน
COPY requirements.txt .

# ติดตั้ง Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# ก๊อปโค้ดทั้งหมด
COPY . .

# สร้าง user ปกติ (HF Spaces ต้องการ — ไม่ใช่ root)
RUN useradd -m -u 1000 user && chown -R user:user /app
USER user

# Streamlit ใช้ home directory ของ user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH \
    PYTHONUNBUFFERED=1 \
    STREAMLIT_SERVER_PORT=7860 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# Expose port
EXPOSE 7860

# Run Streamlit
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=7860", "--server.address=0.0.0.0"]
