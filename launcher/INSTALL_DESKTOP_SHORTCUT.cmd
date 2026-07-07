@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_desktop_shortcut.ps1"
if errorlevel 1 (
  echo.
  echo Не удалось создать ярлык.
  pause
)
