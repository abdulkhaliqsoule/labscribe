@echo off
REM ===========================================================================
REM  LabScribe - installer
REM
REM  Copies LabScribe.exe into this user's programs folder and puts shortcuts
REM  on the Desktop and in the Start menu. Nothing else is needed: LabScribe
REM  carries everything it uses inside itself, so there is no Python and no
REM  other software to install first.
REM
REM  To install:   double-click this file.
REM  To remove:    delete the two shortcuts, then delete the folder shown below.
REM ===========================================================================

setlocal
set "TARGET=%LOCALAPPDATA%\Programs\LabScribe"
set "SOURCE=%~dp0LabScribe.exe"

echo.
echo   LabScribe - installing
echo   ----------------------
echo.

if not exist "%SOURCE%" (
    echo   ERROR: LabScribe.exe was not found next to this installer.
    echo   Keep both files in the same folder and run this again.
    echo.
    pause
    exit /b 1
)

echo   Copying to %TARGET%
if not exist "%TARGET%" mkdir "%TARGET%" >nul 2>&1
copy /y "%SOURCE%" "%TARGET%\LabScribe.exe" >nul
if errorlevel 1 (
    echo   ERROR: could not copy the program. Is LabScribe already running?
    echo.
    pause
    exit /b 1
)

REM Carry across a replacement logo, if one is sitting beside the installer
if exist "%~dp0lab_logo.png" copy /y "%~dp0lab_logo.png" "%TARGET%\" >nul

echo   Creating shortcuts

set "VBS=%TEMP%\labscribe_shortcut.vbs"
> "%VBS%" echo Set oWS = WScript.CreateObject("WScript.Shell")
>>"%VBS%" echo sDesk = oWS.SpecialFolders("Desktop") ^& "\LabScribe.lnk"
>>"%VBS%" echo Set oLink = oWS.CreateShortcut(sDesk)
>>"%VBS%" echo oLink.TargetPath = "%TARGET%\LabScribe.exe"
>>"%VBS%" echo oLink.WorkingDirectory = "%TARGET%"
>>"%VBS%" echo oLink.Description = "LabScribe - clinical chemistry reporting"
>>"%VBS%" echo oLink.Save
>>"%VBS%" echo sMenu = oWS.SpecialFolders("Programs") ^& "\LabScribe.lnk"
>>"%VBS%" echo Set oLink2 = oWS.CreateShortcut(sMenu)
>>"%VBS%" echo oLink2.TargetPath = "%TARGET%\LabScribe.exe"
>>"%VBS%" echo oLink2.WorkingDirectory = "%TARGET%"
>>"%VBS%" echo oLink2.Description = "LabScribe - clinical chemistry reporting"
>>"%VBS%" echo oLink2.Save
cscript //nologo "%VBS%" >nul 2>&1
del "%VBS%" >nul 2>&1

echo.
echo   Done. LabScribe is on your Desktop and in the Start menu.
echo.
echo   The first time you open it, Windows may show a blue
echo   "Windows protected your PC" box, because the program is not
echo   signed with a paid certificate. Click "More info", then
echo   "Run anyway". It will not ask again.
echo.
pause
