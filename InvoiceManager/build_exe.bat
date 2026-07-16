@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "FinishInvoiceManager" --collect-all reportlab --collect-all cryptography app.py
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "FinishLicenseAdmin" --collect-all cryptography license_admin.py
echo.
echo Client EXE: dist\FinishInvoiceManager.exe
echo Admin EXE:  dist\FinishLicenseAdmin.exe
pause
