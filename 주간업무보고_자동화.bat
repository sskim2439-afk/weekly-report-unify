@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
