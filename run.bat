@echo off
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo Something went wrong. Try running install.bat first.
    pause
)
