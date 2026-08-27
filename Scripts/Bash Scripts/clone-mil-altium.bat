@echo off
setlocal

set "REPO_URL=https://github.com/uf-mil-electrical/MIL-Altium.git"
set "DEST=C:\MIL-Altium"

where git >nul 2>nul
if errorlevel 1 (
    echo Git is not installed. Please install it from https://git-scm.com/download/win and re-run this script.
    exit /b 1
)

if exist "%DEST%\" (
    echo %DEST% already exists. Aborting to avoid overwriting or duplicating files.
    exit /b 1
)

echo Cloning into %DEST%...
git clone "%REPO_URL%" "%DEST%"

echo Done.
endlocal
