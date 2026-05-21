@echo off
echo Cek proses yang menggunakan port 5000...
netstat -ano | findstr :5000
echo.
pause
