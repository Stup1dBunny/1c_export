@echo off
setlocal
chcp 65001 >nul
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" -X utf8 "%~dp0nightly.py" --show-report %*
) else (
  "%~dp0..\work\.venv\Scripts\python.exe" -X utf8 "%~dp0nightly.py" --show-report %*
)
set "ExportResult=%ERRORLEVEL%"
echo.
echo Exit code: %ExportResult% (0 = success).
echo Reports folder: %~dp0state
pause
exit /b %ExportResult%
