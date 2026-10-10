@echo off
rem Builds Pursuit.exe (one file, own icon) into this folder.
rem Needs Python 3.10+ with the packages from requirements.txt.
cd /d "%~dp0"
python -m pip install -r requirements.txt
taskkill /F /IM Pursuit.exe >nul 2>&1
python -m PyInstaller --noconfirm --onefile --windowed --name Pursuit --icon "%~dp0pursuit.ico" ^
  --hidden-import win32com.client --hidden-import pythoncom --hidden-import phone_server ^
  --distpath "%~dp0." --workpath "%TEMP%\pursuit_build" --specpath "%TEMP%\pursuit_build" pursuit.py
echo.
if exist Pursuit.exe (echo Built: %~dp0Pursuit.exe) else (echo Build failed - see messages above.)
pause
