@echo off
cd /d "%~dp0"

echo Membuat virtual environment (venv)...
python -m venv venv

echo Mengaktifkan venv dan menginstall dependencies...
call venv\Scripts\activate
pip install -r requirements.txt

echo.
echo ========================================================
echo Setup selesai! Venv telah siap digunakan.
echo ========================================================
