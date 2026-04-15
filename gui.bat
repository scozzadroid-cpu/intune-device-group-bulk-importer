@echo off
REM Intune Device Group Bulk Importer — Batch wrapper
REM Launches gui.ps1 with proper PowerShell configuration

setlocal enabledelayedexpansion
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "gui.ps1"
exit /b %ERRORLEVEL%
