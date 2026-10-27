@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
    echo Khong tim thay Python Launcher ^(py^^). Hay cai Python va bat tuy chon Add Python to PATH.
    pause
    exit /b 1
)
py -3 -m venv .venv
if errorlevel 1 goto failed
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto failed
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo.
echo Cai dat hoan tat. Chay run_windows.bat de mo ung dung.
pause
exit /b 0
:failed
echo.
echo Cai dat that bai. Hay kiem tra ket noi Internet va phien ban Python.
pause
exit /b 1
