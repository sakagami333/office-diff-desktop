@echo off
rem Drop one or two .docx/.pptx files onto this icon (or double-click to pick files).
rem The older file (by modified time) is treated as "before".
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0OfficeDiff.ps1" %*
echo.
pause
