@echo off
setlocal
:: ============================================================
::  Office to PDF Converter - add "Convert to PDF" to the right-click menu
:: ============================================================
::  - Works for .doc .docx .xls .xlsx .ppt .pptx and for folders
::  - Installs for the current user only (no admin needed)
::  - Run uninstall_context_menu.bat to remove it again
::  - If you move the office-to-pdf folder, run this file again
:: ============================================================

set "HERE=%~dp0"
set "SCRIPT=%HERE%convert.py"
set "ICON=%HERE%pdf.ico"

if not exist "%SCRIPT%" (
    echo [ERROR] convert.py was not found next to this file.
    pause
    exit /b 1
)

:: ---- Find the real python.exe (not the Microsoft Store shortcut) ----
set "PY="
for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%P"
if not defined PY (
    for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%P"
)
if not defined PY (
    echo [ERROR] Python was not found.
    echo Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup.
    pause
    exit /b 1
)

echo Python : %PY%
echo Script : %SCRIPT%
echo.

:: ---- Remove the old menu entry from earlier versions ----
for %%E in (.doc .docx .xls .xlsx .ppt .pptx) do reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\REAConvertToPDF" /f >nul 2>&1
reg delete "HKCU\Software\Classes\Directory\shell\REAConvertToPDF" /f >nul 2>&1

:: ---- Register the menu item for each file type ----
for %%E in (.doc .docx .xls .xlsx .ppt .pptx) do call :addkey "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\OfficeToPDF" %%E

:: ---- And for folders (converts every Office file inside) ----
call :addkey "HKCU\Software\Classes\Directory\shell\OfficeToPDF" folders

echo.
echo ============================================================
echo  Done! Right-click any Word, Excel or PowerPoint file and
echo  choose "Convert to PDF".
echo.
echo  Windows 11: the item is under "Show more options"
echo  (or hold Shift while you right-click).
echo ============================================================
pause
exit /b 0

:addkey
reg add "%~1" /ve /d "Convert to PDF" /f >nul
reg add "%~1" /v Icon /d "%ICON%" /f >nul
reg add "%~1\command" /ve /d "\"%PY%\" \"%SCRIPT%\" --from-menu \"%%1\"" /f >nul
if errorlevel 1 (
    echo [FAILED] %2
) else (
    echo [OK] %2
)
exit /b 0
