@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
    echo Chua tao moi truong ao. Hay chay setup_windows.bat truoc.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
python app.py
pause
