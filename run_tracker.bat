@echo off
setlocal
cd /d "%~dp0"

REM --- Check for Administrative Privileges ---
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [TOBII] Running with Administrative privileges.
    goto :run_python
)

REM --- Not Admin: Try Silent Run via Task Scheduler ---
echo [TOBII] Not Admin. Attempting silent launch (UAC bypass)...
schtasks /run /tn "TobiiMouseController" >nul 2>&1
if %errorLevel% == 0 (
    echo [SUCCESS] Silent task triggered. This window will close.
    timeout /t 2 >nul
    exit /b 0
)

echo [WARNING] Silent task 'TobiiMouseController' not found.
echo [TIP] Run 'setup_silent_run.bat' once to enable automatic UAC bypass.
echo.
echo Running in current console (may show UAC prompt if needed)...

:run_python
echo [TOBII] Starting Tracker script...
set "PYEXE=%~1"
if "%PYEXE%"=="" if exist python_path.txt set /p PYEXE=<python_path.txt
if "%PYEXE%"=="" set "PYEXE=%USERPROFILE%\AppData\Local\Programs\Python\Python310-32\python.exe"
echo [TOBII] Python Path: "%PYEXE%"
if not exist "%PYEXE%" (
    echo [ERROR] Python 3.10-32 not found at "%PYEXE%"
    echo [TIP] Please install 32-bit Python 3.10 or run 'setup_silent_run.bat' as Admin.
    pause
    exit /b 1
)
"%PYEXE%" tobii_native.py
if errorlevel 1 pause
