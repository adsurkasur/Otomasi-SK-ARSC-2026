@echo off
REM Pindah ke direktori di mana script berada
cd /d "%~dp0"

REM Aktifkan virtual environment (buat jika belum ada)
if not exist venv\Scripts\activate.bat (
    call setup_venv.bat
)
call venv\Scripts\activate

REM Periksa apakah ada argumen yang diberikan
if "%~1"=="" (
    echo.
    echo Tidak ada nama yang dimasukkan!
    echo.
    echo Cara Penggunaan:
    echo 1. Drag-and-drop file .txt berisi daftar nama ke file .bat ini
    echo 2. ATAU jalankan dari CMD: run_generate.bat "Nama 1" "Nama 2"
    echo 3. ATAU jalankan dari CMD: run_generate.bat --all
    echo.
    echo Tekan tombol apa saja untuk keluar...
    pause >nul
    exit /b
)

REM Jika argumen pertama adalah file (drag and drop file .txt)
if exist "%~1" (
    echo Memproses daftar nama dari file: %~1
    python generate_sk.py --file "%~1"
) else (
    REM Jika argumen berupa nama-nama dari CMD
    python generate_sk.py %*
)

echo.
pause
