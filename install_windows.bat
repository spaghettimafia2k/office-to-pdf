@echo off
:: ===================================================
::  Office to PDF Converter - Windows setup (run once)
:: ===================================================

echo ===================================================
echo  Office to PDF Converter - Windows Setup
echo ===================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not in PATH.
        echo Download it from https://www.python.org/downloads/
        echo and tick "Add python.exe to PATH" during setup.
        pause
        exit /b 1
    )
    set "PYCMD=py -3"
) else (
    set "PYCMD=python"
)
echo [OK] Python found.

echo.
echo Installing pywin32 (lets the converter use Microsoft Office)...
%PYCMD% -m pip install --upgrade pywin32
if errorlevel 1 (
    echo [WARNING] pywin32 install failed. LibreOffice will be used instead.
) else (
    echo [OK] pywin32 installed.
)

echo.
echo Adding "Convert to PDF" to the right-click menu...
call "%~dp0install_context_menu.bat"
