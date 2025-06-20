@echo off

setlocal

:: Define paths

set "USERPROFILE=%USERPROFILE%"

set "DOWNLOADS=%USERPROFILE%\Downloads"

set "OUTLOOK_FOLDER=%APPDATA%\Microsoft\Outlook"

set "VBAPROJECT=%DOWNLOADS%\VbaProject.otm"

set "ZIPFILE=%DOWNLOADS%\xpdf-tools-win-4.05.zip"

set "DEST_FOLDER=%USERPROFILE%\Documents\PDFTools"

:: Move VbaProject.otm

if exist "%VBAPROJECT%" (

    echo Moving VbaProject.otm to Outlook folder...

    move /Y "%VBAPROJECT%" "%OUTLOOK_FOLDER%"

) else (

    echo VbaProject.otm not found in Downloads.

)

:: Create destination folder if not exists

if not exist "%DEST_FOLDER%" (

    mkdir "%DEST_FOLDER%"

)

:: Unzip using 7-Zip

if exist "%ZIPFILE%" (

    echo Extracting xpdf-tools to Documents\PDFTools...

    "C:\Program Files\7-Zip\7z.exe" x "%ZIPFILE%" -o"%DEST_FOLDER%" -y

) else (

    echo xpdf-tools-win-4.05.zip not found in Downloads.

)

echo Done.

pause
 