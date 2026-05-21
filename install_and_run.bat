@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo Excel Delete Insert SQL Generator - Setup
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PY=python"
    ) else (
        echo ERROR: Python belum terinstall.
        echo Install Python 3.11 atau 3.12 dulu dari https://www.python.org/downloads/
        echo Pastikan centang "Add python.exe to PATH" saat install.
        pause
        exit /b 1
    )
)

echo Membuat virtual environment...
%PY% -m venv .venv
if %errorlevel% neq 0 (
    echo ERROR: Gagal membuat virtual environment.
    pause
    exit /b 1
)

echo Mengaktifkan virtual environment...
call ".venv\Scripts\activate.bat"

echo Upgrade pip...
python -m pip install --upgrade pip

echo Install dependency...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Gagal install dependency.
    pause
    exit /b 1
)

echo.
echo ============================================
echo Aplikasi akan berjalan di:
echo http://localhost:5000
echo ============================================
echo Jangan tutup window ini selama aplikasi dipakai.
echo.

python sql_delete_insert_webapp_v7.py

pause
