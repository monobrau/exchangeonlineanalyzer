@echo off
title Microsoft 365 Management Tool
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "launcher.ps1"
pause
