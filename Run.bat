@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File ".\Inventory Health Checker.ps1" -ExportHtml -OpenReport
exit
