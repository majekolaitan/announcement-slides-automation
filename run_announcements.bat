@echo off
title SDA Kubwa PPT System
setlocal enabledelayedexpansion
color 0A

echo ============================================
echo   SDA KUBWA ANNOUNCEMENT SYSTEM
echo ============================================
echo.

:: 1. Try to find Python (checks 'python' then 'py')
set PY_CMD=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PY_CMD=python
) else (
    py --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PY_CMD=py
    )
)

if not defined PY_CMD (
    color 0C
    echo [!] ERROR: Python was NOT found on this system.
    echo 1. Install Python from https://www.python.org/
    echo 2. IMPORTANT: Check the box "Add Python to PATH" during installation.
    echo.
    pause
    exit
)

:: 2. Check if the JSON file exists before starting
if not exist "announcements.json" (
    color 0C
    echo [!] ERROR: 'announcements.json' not found!
    echo Please make sure the file is named correctly and is in this folder.
    echo.
    pause
    exit
)

:: 3. Auto-install requirements
echo [1/3] Checking dependencies...
%PY_CMD% -c "import pptx" 2>nul
if %errorlevel% neq 0 (
    echo Installing required libraries...
    %PY_CMD% -m pip install python-pptx --quiet --disable-pip-version-check
)

cls

:: 4. Run the generator
echo [2/3] Generating PowerPoint...
%PY_CMD% build_slides.py

:: 5. Error Catching
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [!] GENERATION FAILED. 
    echo Check the error message above. Likely a typo in your JSON file.
    echo.
    pause
    exit
)

:: 6. Find and Open the PPTX
echo [3/3] Opening presentation...
set "outfile=SDA_Kubwa_Announcements.pptx"

if exist "%outfile%" (
    echo [SUCCESS] Opening: %outfile%
    start "" "%outfile%"
    echo.
    echo Process Complete!
) else (
    color 0C
    echo [!] Error: PowerPoint was not created.
)

:: Keep window open for 5 seconds then close, or press key to close
echo.
echo Window will close automatically in 5 seconds...
timeout /t 5