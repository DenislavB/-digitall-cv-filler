@echo off
title CV Filler — Build Installer
echo.
echo ============================================
echo  CV Filler ^| DIGITALL Format — Build Tool
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python from python.org first.
    pause & exit /b 1
)

:: Install/upgrade dependencies
echo [1/3] Installing dependencies...
pip install --quiet --upgrade pdfplumber python-docx lxml pyinstaller
if errorlevel 1 (
    echo ERROR: pip install failed. Check your internet connection.
    pause & exit /b 1
)
echo       Done.
echo.

:: Build the executable
echo [2/3] Building executable (this takes ~60 seconds)...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "CV-Filler-DIGITALL" ^
    --add-data "template.docx;." ^
    --hidden-import "pdfplumber" ^
    --hidden-import "pdfminer" ^
    --hidden-import "pdfminer.high_level" ^
    --hidden-import "pdfminer.layout" ^
    --hidden-import "docx" ^
    --hidden-import "lxml" ^
    --hidden-import "lxml.etree" ^
    --hidden-import "PIL" ^
    --collect-submodules pdfplumber ^
    --collect-submodules pdfminer ^
    --noconfirm ^
    app.py

if errorlevel 1 (
    echo ERROR: Build failed. See output above for details.
    pause & exit /b 1
)
echo       Done.
echo.

:: Package output
echo [3/3] Packaging...
if not exist "release" mkdir release
copy /Y "dist\CV-Filler-DIGITALL.exe" "release\CV-Filler-DIGITALL.exe" >nul
copy /Y "template.docx" "release\template.docx" >nul

echo.
echo ============================================
echo  BUILD COMPLETE!
echo ============================================
echo.
echo  Files in the  release\  folder:
echo    CV-Filler-DIGITALL.exe  ^<-- send this to colleagues
echo    template.docx           ^<-- must stay next to the .exe
echo.
echo  Both files must be kept in the same folder.
echo  No Python or install needed on the target PC.
echo.
pause
