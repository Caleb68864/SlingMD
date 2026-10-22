@echo off
rem Double-click to run Trust-SlingMD.ps1 without changing the PowerShell execution policy.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Trust-SlingMD.ps1" %*
