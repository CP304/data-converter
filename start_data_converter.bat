@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" data_converter_gui.py
  exit /b %errorlevel%
)
where python >nul 2>nul
if %errorlevel%==0 (
  python data_converter_gui.py
  exit /b %errorlevel%
)
where py >nul 2>nul
if %errorlevel%==0 (
  py data_converter_gui.py
  exit /b %errorlevel%
)
echo Kein Python gefunden. Bitte Python installieren oder eine .venv im Projektordner anlegen.
pause
