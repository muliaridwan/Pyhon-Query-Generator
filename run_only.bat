@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment belum ada.
    echo Jalankan install_and_run.bat dulu.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

echo ============================================
echo Running Excel Delete Insert SQL Generator
echo URL: http://localhost:5000
echo ============================================
echo Jangan tutup window ini selama aplikasi dipakai.
echo.

python sql_delete_insert_webapp_v7.py

pause
