@echo off
setlocal
title Tidying the LabScribe folder

REM ===========================================================================
REM  Moves superseded builds into "Older versions". Nothing is deleted, and
REM  the current LabScribe.exe and LabScribe.py are never touched.
REM
REM  Safe to run as often as you like. If there is nothing to move it says so
REM  and stops.
REM
REM  (An earlier version of this file moved LabScribe.exe aside on the
REM  assumption it was still 2.3. Now that LabScribe.exe IS 2.4, doing that
REM  would file the current program away as an old one — so it no longer
REM  touches it at all.)
REM ===========================================================================

cd /d "%~dp0"

echo.
echo   Tidying the LabScribe folder
echo   ---------------------------
echo.

set OLD=Older versions
if not exist "%OLD%" mkdir "%OLD%"

set MOVED=0

REM -- Builds that carry a version number are by definition not the current one -
for %%F in (LabScribe_2.1.exe LabScribe_2.2.exe LabScribe_2.3.exe
            LabScribe_2.4.exe LabScribe_2.5.exe) do (
    if exist "%%F" (
        move /y "%%F" "%OLD%\%%F" >nul
        if errorlevel 1 (
            echo   could not move %%F  - it is probably open. Close it and
            echo                         run this again.
        ) else (
            echo   moved  %%F
            set /a MOVED+=1
        )
    )
)

REM -- Sources of superseded versions -------------------------------------------
for %%F in (LabScribe_2.1.py LabScribe_2.2.py LabScribe_2.3.py
            LabScribe_2.4.py LabScribe_2.5.py) do (
    if exist "%%F" (
        move /y "%%F" "%OLD%\%%F" >nul
        if errorlevel 1 (
            echo   could not move %%F
        ) else (
            echo   moved  %%F
            set /a MOVED+=1
        )
    )
)

REM -- Build scripts that only knew about one version ---------------------------
for %%F in ("Build LabScribe 2.4.bat" "Build LabScribe 2.5.bat") do (
    if exist %%F (
        move /y %%F "%OLD%\%%~nxF" >nul
        if not errorlevel 1 (
            echo   moved  %%~nxF
            set /a MOVED+=1
        )
    )
)

echo.
if "%MOVED%"=="0" (
    echo   Nothing to move - the folder is already tidy.
) else (
    echo   Done. %MOVED% item^(s^) moved into "%OLD%".
)
echo.
echo   The folder should now hold:
echo       LabScribe.exe       the program to use
echo       LabScribe.py        its source
echo       labscribe.ico       its icon
echo       Build LabScribe.bat builds the exe from the source
echo       Older versions\     everything superseded, kept in case
echo.
echo   To put an old build back, drag it out of "Older versions".
echo.
pause
