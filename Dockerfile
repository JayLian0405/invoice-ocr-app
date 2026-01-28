# 使用官方 Python 輕量版
FROM python:3.9-slim

# 設定工作目錄
WORKDIR /app

# 安裝系統套件 (給 PyMuPDF 用)
RUN apt-get update && apt-get install -y libgl1 libglib2.0-0

# --- 關鍵修正 ---
# 我們強制指定 "numpy<2.0.0"，避免新版衝突導致閃退
# 同時直接在這裡安裝所有套件，繞過 requirements.txt 的讀取問題
RUN pip install "numpy<2.0.0" flask gunicorn google-generativeai python-dotenv requests google-api-python-client google-auth-httplib2 google-auth-oauthlib pymupdf pandas openpyxl xlwt
# 複製所有程式碼
COPY . .

# 設定環境變數 (讓 Log 可以馬上印出來，方便除錯)
ENV PYTHONUNBUFFERED=1

# 啟動指令
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app1:app