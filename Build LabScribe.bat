@echo off
setlocal enabledelayedexpansion
title Building LabScribe

REM ===========================================================================
REM  Builds LabScribe.exe from LabScribe.py, with labscribe.ico as its icon.
REM
REM  The exe is only put in place once it has actually been built. Whatever
REM  LabScribe.exe was there before is moved into "Older versions" first,
REM  under a dated name, so nothing is ever written over.
REM
REM  Needs Python on this PC. Everything else it fetches itself.
REM ===========================================================================

cd /d "%~dp0"

echo.
echo   Building LabScribe
echo   ------------------
echo.

REM -- Nothing can be replaced while it is running ------------------------------
tasklist /fi "imagename eq LabScribe.exe" 2>nul | find /i "LabScribe.exe" >nul
if not errorlevel 1 (
    echo   LabScribe is open. Close it and run this again.
    echo.
    pause
    exit /b 1
)

REM -- Find Python -------------------------------------------------------------
set PY=
py -3 --version >nul 2>&1 && set PY=py -3
if not defined PY ( python --version >nul 2>&1 && set PY=python )

if not defined PY (
    echo   Python is not on this PC.
    echo.
    echo   Install it from https://www.python.org/downloads/
    echo   Tick "Add python.exe to PATH" on the first screen of the installer,
    echo   then run this file again.
    echo.
    pause
    exit /b 1
)

for /f "delims=" %%v in ('%PY% --version 2^>^&1') do echo   Using %%v

if not exist "LabScribe.py" (
    echo.
    echo   LabScribe.py is not in this folder.
    echo.
    pause
    exit /b 1
)

REM -- Everything the program and the builder need ------------------------------
echo.
echo   Fetching what it needs. A few minutes the first time, seconds after.
echo.
%PY% -m pip install --upgrade pip >nul 2>&1
%PY% -m pip install --upgrade pyinstaller reportlab pdfplumber tkinterdnd2 pillow
if errorlevel 1 (
    echo.
    echo   Those could not be fetched. Check this PC is on the internet.
    echo.
    pause
    exit /b 1
)

REM -- The program icon ---------------------------------------------------------
set ICON=
set ICONDATA=
if exist "labscribe.ico" (
    set ICON=--icon "labscribe.ico"
    set ICONDATA=--add-data "labscribe.ico;."
    echo   Using labscribe.ico for the program icon.
) else (
    echo   labscribe.ico is not here - building without an icon.
)

REM -- Build --------------------------------------------------------------------
echo.
echo   Building. Leave it alone until it finishes.
echo.

%PY% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name LabScribe ^
    --collect-all tkinterdnd2 ^
    --collect-all pypdfium2 ^
    --collect-data reportlab ^
    --collect-data pdfminer ^
    --hidden-import pdfplumber ^
    --hidden-import tkinterdnd2 ^
    --hidden-import PIL ^
    --noconfirm ^
    --clean ^
    %ICON% ^
    %ICONDATA% ^
    "LabScribe.py"

if errorlevel 1 (
    echo.
    echo   The build failed. The lines above say why - copy them and send
    echo   them on. Nothing in this folder has been changed.
    echo.
    pause
    exit /b 1
)

if not exist "dist\LabScribe.exe" (
    echo.
    echo   The build finished but no exe came out of it. Nothing in this
    echo   folder has been changed.
    echo.
    pause
    exit /b 1
)

REM -- Put it in place, keeping whatever was there ------------------------------
set OLD=Older versions
if not exist "%OLD%" mkdir "%OLD%"

if exist "LabScribe.exe" (
    set STAMP=%DATE:/=-%
    set STAMP=!STAMP: =_!
    move /y "LabScribe.exe" "%OLD%\LabScribe_before_!STAMP!.exe" >nul
    if errorlevel 1 (
        echo   The previous LabScribe.exe could not be moved aside. The new
        echo   build is waiting in the "dist" folder.
        echo.
        pause
        exit /b 1
    )
    echo   Previous build moved into "%OLD%".
)

move /y "dist\LabScribe.exe" "LabScribe.exe" >nul
if errorlevel 1 (
    echo   The new exe could not be moved out of "dist". It is still there.
    echo.
    pause
    exit /b 1
)

REM -- Any superseded 2.4 build left lying about --------------------------------
if exist "LabScribe_2.4.exe" (
    move /y "LabScribe_2.4.exe" "%OLD%\LabScribe_2.4_early.exe" >nul
    if not errorlevel 1 echo   Earlier 2.4 build moved into "%OLD%".
)

rmdir /s /q build >nul 2>&1
rmdir /s /q dist  >nul 2>&1
del /q "LabScribe.spec" >nul 2>&1

echo.
echo   Done.
echo.
echo   LabScribe.exe is in this folder, with the new icon.
echo   Double-click it and put a couple of real analyser PDFs through it.
echo.
echo   Windows will say "Windows protected your PC" the first time, because
echo   the program is not signed. Click "More info", then "Run anyway".
echo.
pause
