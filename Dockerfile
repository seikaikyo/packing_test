# 使用 Python 作為基礎映像
FROM python:3.9-slim

# 安裝 tkinter 及其他所需套件
RUN apt-get update && apt-get install -y \
    python3-tk \
    x11-apps \
    xvfb \
    && rm -rf /var/lib/apt/lists/*

# 設定工作目錄
WORKDIR /app

# 把目前的程式碼複製到 Docker 裡的 /app 目錄
COPY . /app

# 安裝 Python 依賴的套件
RUN pip install --no-cache-dir -r requirements.txt

# 設定虛擬顯示環境
ENV DISPLAY=:99

# 運行虛擬顯示器，然後運行你的程式
CMD ["sh", "-c", "Xvfb :99 -screen 0 1024x768x16 & python your_script.py"]
