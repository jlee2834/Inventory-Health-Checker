@echo off
cd /d "%~dp0"
echo Running from: %cd%
echo Starting script...
powershell.exe -NoExit -ExecutionPolicy Bypass -File ".\InventoryHealthChecker.ps1" -ExportHtml -OpenReport
pause
