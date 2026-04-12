@echo off
setlocal
cd /d "%~dp0"

python -m PyInstaller --noconfirm --windowed --onefile --icon app_icon.ico --name "Citizen StarString Updater Helper" updater_helper.py
if errorlevel 1 exit /b 1

python -m PyInstaller --noconfirm --windowed --onefile --icon app_icon.ico --add-data "app_icon.png;." --add-data "app_icon.ico;." --add-data "dist\Citizen StarString Updater Helper.exe;." --name "Citizen StarString Helper" starstrings_updater.py
if exist "dist\Citizen StarString Helper.exe" copy /y "dist\Citizen StarString Helper.exe" "Citizen StarString Helper.exe"
endlocal
