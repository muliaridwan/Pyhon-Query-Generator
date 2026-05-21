DEPLOY LOCAL - Excel Delete Insert SQL Generator

ISI PAKET
1. sql_delete_insert_webapp_v7.py
   Aplikasi FastAPI untuk upload Excel dan generate SQL delete + insert.

2. requirements.txt
   Dependency Python.

3. install_and_run.bat
   Script setup pertama kali:
   - membuat virtual environment .venv
   - install dependency
   - menjalankan aplikasi di http://localhost:5000

4. run_only.bat
   Script menjalankan aplikasi setelah setup pertama selesai.

5. check_port_5000.bat
   Cek apakah port 5000 sedang dipakai.

6. stop_port_5000.bat
   Mematikan proses yang sedang LISTENING di port 5000.


CARA PAKAI PERTAMA KALI
1. Extract ZIP ke folder lokal, contoh:
   C:\Tools\excel-sql-generator

2. Pastikan Python 3.11 atau 3.12 sudah terinstall.
   Saat install Python, centang:
   Add python.exe to PATH

3. Double-click:
   install_and_run.bat

4. Buka browser:
   http://localhost:5000


CARA RUN BERIKUTNYA
Double-click:
run_only.bat


JIKA PORT 5000 BENTROK
1. Jalankan:
   check_port_5000.bat

2. Jika ingin mematikan proses di port 5000:
   stop_port_5000.bat

3. Jalankan ulang:
   run_only.bat


CATATAN PENTING
- Ini bukan aplikasi ASP.NET/IIS native.
- Ini aplikasi Python FastAPI yang berjalan melalui Uvicorn.
- Window command prompt jangan ditutup selama aplikasi dipakai.
- Kalau ingin dibuat auto-start Windows Service, lebih aman pakai NSSM atau Task Scheduler.
