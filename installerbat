@echo off

setlocal

:: Kill Outlook if running

tasklist /FI "IMAGENAME eq OUTLOOK.EXE" | find /I "OUTLOOK.EXE" >nul

if %errorlevel%==0 (

    echo Closing Outlook...

    taskkill /IM OUTLOOK.EXE /F >nul

    timeout /t 2 >nul

)

:: Paths

set "USERPROFILE=%USERPROFILE%"

set "DOWNLOADS=%USERPROFILE%\Downloads"

set "OUTLOOK_FOLDER=%APPDATA%\Microsoft\Outlook"

set "VBAPROJECT=%DOWNLOADS%\VbaProject.otm"

set "ZIPFILE=%DOWNLOADS%\xpdf-tools-win-4.05.zip"

set "DOCS_FOLDER=%USERPROFILE%\Documents"

set "DEST_FOLDER=%DOCS_FOLDER%\PDFTools"

:: Move .otm

if exist "%VBAPROJECT%" (

    echo Copying VbaProject.otm to Outlook folder...

    copy /Y "%VBAPROJECT%" "%OUTLOOK_FOLDER%\VbaProject.otm" >nul

    echo .otm copied.

) else (

    echo VbaProject.otm not found in Downloads.

)

:: Unzip and rename

if exist "%ZIPFILE%" (

    echo Extracting xpdf-tools to Documents...

    "C:\Program Files\7-Zip\7z.exe" x "%ZIPFILE%" -o"%DOCS_FOLDER%" -y >nul

    if exist "%DOCS_FOLDER%\xpdf-tools-win-4.05" (

        if exist "%DEST_FOLDER%" (

            echo Removing old PDFTools folder...

            rmdir /s /q "%DEST_FOLDER%"

        )

        echo Renaming xpdf-tools-win-4.05 to PDFTools...

        ren "%DOCS_FOLDER%\xpdf-tools-win-4.05" "PDFTools"

    )

    echo Extraction complete.

) else (

    echo xpdf-tools-win-4.05.zip not found in Downloads.

)

echo Done.

pause
 
