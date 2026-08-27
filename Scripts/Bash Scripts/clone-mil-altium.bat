@echo off
setlocal

set "REPO_URL=https://github.com/uf-mil-electrical/MIL-Altium.git"
set "DEST=C:\MIL-Altium"


echo **********clone-mil-altium.bat**********
echo ^> Hello! This script will now clone the MIL-Altium repo to C:/MIL-Altium.

where git >nul 2>nul
if errorlevel 1 (
    echo ^> Git is not installed. Please install it from https://git-scm.com/download/win and re-run this script.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

if exist "%DEST%\" (
    echo ^> %DEST% already exists. Aborting to avoid overwriting or duplicating files.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

git clone "%REPO_URL%" "%DEST%"
if errorlevel 1 (
    echo Clone failed. See the error above.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

echo ^> Done.
timeout /t 2 /nobreak >nul
endlocal
