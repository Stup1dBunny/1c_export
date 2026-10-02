@echo off
call "%~dp0automation\run_reports.cmd" %*
exit /b %ERRORLEVEL%
