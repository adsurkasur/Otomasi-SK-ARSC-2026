@echo off
REM Pindah ke direktori di mana script berada
cd /d "%~dp0"

REM Aktifkan virtual environment (buat jika belum ada)
if not exist venv\Scripts\activate.bat (
    call setup_venv.bat
)
call venv\Scripts\activate

echo ===================================================
echo Memulai Program Konversi DOCX ke PDF...
echo ===================================================
echo.
python convert_to_pdf.py %*

echo.
echo ===================================================
echo Selesai!
echo ===================================================
pause
