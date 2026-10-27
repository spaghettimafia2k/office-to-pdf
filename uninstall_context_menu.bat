@echo off
:: ============================================================
::  Office to PDF Converter - remove "Convert to PDF" from the right-click menu
:: ============================================================

for %%E in (.doc .docx .xls .xlsx .ppt .pptx) do (
    reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\OfficeToPDF" /f >nul 2>&1
    reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\REAConvertToPDF" /f >nul 2>&1
)
reg delete "HKCU\Software\Classes\Directory\shell\OfficeToPDF" /f >nul 2>&1
reg delete "HKCU\Software\Classes\Directory\shell\REAConvertToPDF" /f >nul 2>&1

echo "Convert to PDF" has been removed from the right-click menu.
pause
