@echo off
chcp 65001 >nul
title مطابقة المشتريات والمبيعات
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python غير مثبت على هذا الجهاز. ثبّته من python.org ثم أعد التشغيل.
    pause
    exit /b 1
)

python -c "import fastapi, uvicorn, openpyxl, requests, multipart, psycopg2, dotenv" >nul 2>nul
if errorlevel 1 (
    echo جارٍ تثبيت المتطلبات لأول مرة...
    pip install -r requirements.txt
)

echo.
echo تشغيل الأداة على  http://127.0.0.1:8077  ... أغلق هذه النافذة لإيقافها.
start "" "http://127.0.0.1:8077"
python -m uvicorn app.main:app --host 127.0.0.1 --port 8077
