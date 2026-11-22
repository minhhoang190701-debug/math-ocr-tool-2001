FROM python:3.11-slim

# Cài đặt các gói hệ thống cần thiết
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy file requirements và cài đặt
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Cấu hình User 1000 chuẩn Hugging Face Spaces (tránh lỗi phân quyền file)
RUN useradd -m -u 1000 user
USER user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH \
    PORT=7860

WORKDIR $HOME/app

# Copy toàn bộ mã nguồn và cấp quyền cho user
COPY --chown=user . $HOME/app

# Tạo sẵn thư mục lưu file tạm
RUN mkdir -p $HOME/app/temp_output

EXPOSE 7860

CMD ["python", "app.py"]
