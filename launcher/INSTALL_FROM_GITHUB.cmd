@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL_FROM_GITHUB.ps1"
if errorlevel 1 (
  echo.
  echo Не удалось установить приложение.
  pause
)
