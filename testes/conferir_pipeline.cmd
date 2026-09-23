@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0conferir_pipeline.ps1" %*
