@echo off
rem Same as OfficeDiff.bat, but swaps before/after.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0OfficeDiff.ps1" -Swap %*
echo.
pause
