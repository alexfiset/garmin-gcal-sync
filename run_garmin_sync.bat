@echo off
REM Windows Task Scheduler wrapper for garmin_sync.py
REM Edit PYTHON_EXE if you use a virtualenv.

setlocal
set SCRIPT_DIR=%~dp0
set PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe

cd /d "%SCRIPT_DIR%"
"%PYTHON_EXE%" "%SCRIPT_DIR%garmin_sync.py"
exit /b %ERRORLEVEL%
