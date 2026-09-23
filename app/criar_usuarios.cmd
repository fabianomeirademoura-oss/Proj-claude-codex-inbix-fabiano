@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar_usuarios.ps1" %*
