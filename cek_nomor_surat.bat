@echo off
REM Pindah ke direktori di mana script berada
cd /d "%~dp0"

REM Aktifkan virtual environment (buat jika belum ada)
if not exist venv\Scripts\activate.bat (
    call setup_venv.bat
)
call venv\Scripts\activate

REM Jalankan script cek nomor
python scripts\cek_nomor.py

echo.
pause
