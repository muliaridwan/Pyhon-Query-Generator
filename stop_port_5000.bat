@echo off
setlocal
echo Mencari proses LISTENING di port 5000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
    echo Mematikan PID %%a
    taskkill /PID %%a /F
)
echo Selesai.
pause
