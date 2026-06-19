@echo off
cd /d "%~dp0"

echo PERINGATAN: Direktori venv akan dihapus!
set /p konfirmasi="Lanjutkan? (y/n): "
if /i "%konfirmasi%" NEQ "y" goto cancel

if exist venv (
    echo Menghapus venv...
    rmdir /s /q venv
    echo Berhasil dihapus.
) else (
    echo Direktori venv tidak ditemukan.
)
goto end

:cancel
echo Dibatalkan.

:end
pause
