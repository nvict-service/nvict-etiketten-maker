@echo off
SETLOCAL ENABLEDELAYEDEXPANSION
REM ===================================================
REM NVict Etiketten Maker - Complete Release v3.0
REM PyInstaller + Inno Setup + FTP Upload (WinSCP of PowerShell)
REM ===================================================

cls
echo.
echo  ============================================
echo   NVict Etiketten Maker - Release v3.0
echo  ============================================
echo.

REM ===== PYTHON FILE DETECTIE =====
echo [1/8] Python bestand zoeken...
set PYTHON_FILE=NVict_Etiketten.py
if not exist "%PYTHON_FILE%" (
    echo [!] FOUT: %PYTHON_FILE% niet gevonden!
    pause & exit /b 1
)
echo [OK] %PYTHON_FILE%
echo.

REM ===== VERSIE DETECTIE =====
echo [2/8] Versienummer detecteren...
set VERSION=
for /f "tokens=2 delims== " %%a in ('findstr /C:"APP_VERSION = " "%PYTHON_FILE%"') do set VERSION=%%a
set VERSION=%VERSION:"=%
if "%VERSION%"=="" (
    echo [!] Kon versienummer niet detecteren!
    pause & exit /b 1
)
echo [OK] Versie: %VERSION%
echo.

REM ===== RELEASE NOTES =====
echo [3/8] Release notes...
if not exist "release_notes.txt" (
    echo [i] release_notes.txt niet gevonden - template aanmaken...
    (
        echo • Eén adres op alle etiketten ^(nieuw^)
        echo • Keuze aantal vellen voor één adres
        echo • Verbeterde interface en preview
        echo • Diverse bug fixes en verbeteringen
    ) > release_notes.txt
    echo [OK] Template aangemaakt
    echo.
    set /p "EDIT_NOTES=Open Notepad om te bewerken? (J/N): "
    if /i "!EDIT_NOTES!"=="J" start /wait notepad.exe release_notes.txt
) else (
    echo [OK] release_notes.txt gevonden
    echo.
    echo Huidige inhoud:
    echo  ----------------------------------------
    type release_notes.txt
    echo  ----------------------------------------
    echo.
    set /p "EDIT_NOTES=Open Notepad om te bewerken? (J/N): "
    if /i "!EDIT_NOTES!"=="J" start /wait notepad.exe release_notes.txt
)
echo.

REM ===== PRE-FLIGHT CHECKS =====
echo [4/8] Pre-flight checks...

python --version >nul 2>&1
if errorlevel 1 ( echo [!] Python NIET gevonden & pause & exit /b 1 )
echo [OK] Python

pyinstaller --version >nul 2>&1
if errorlevel 1 ( echo [!] PyInstaller NIET gevonden & pause & exit /b 1 )
echo [OK] PyInstaller

set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if defined ISCC ( echo [OK] Inno Setup 6 ) else ( echo [i] Inno Setup niet gevonden - stap 6 wordt overgeslagen )

REM WinSCP detection
set "WINSCP="
if exist "C:\Program Files (x86)\WinSCP\WinSCP.com" set "WINSCP=C:\Program Files (x86)\WinSCP\WinSCP.com"
if exist "C:\Program Files\WinSCP\WinSCP.com"       set "WINSCP=C:\Program Files\WinSCP\WinSCP.com"
if defined WINSCP (
    echo [OK] WinSCP gevonden - FTP via WinSCP
) else (
    echo [i] WinSCP niet gevonden - FTP via PowerShell fallback
)

if not exist "NVictEtiketten.spec" ( echo [!] NVictEtiketten.spec NIET gevonden! & pause & exit /b 1 )
echo [OK] NVictEtiketten.spec

echo.
echo Alles gereed. Release wordt nu automatisch uitgevoerd...
echo.
timeout /t 3 /nobreak >nul

REM ===== STAP 5: VERSIE UPDATE IN ISS =====
cls
echo.
echo  ============================================
echo   STAP 5: ISS versie bijwerken
echo  ============================================
echo.

set "ISS_FILE=NVict_Etiketten.iss"
if exist "%ISS_FILE%" (
    copy "%ISS_FILE%" "%ISS_FILE%.backup" >nul 2>&1
    powershell -Command "(Get-Content '%ISS_FILE%') -replace '#define AppVersion \".*\"', '#define AppVersion \"%VERSION%\"' | Set-Content '%ISS_FILE%.tmp'"
    move /y "%ISS_FILE%.tmp" "%ISS_FILE%" >nul 2>&1
    echo [OK] Versie %VERSION% in %ISS_FILE% bijgewerkt
) else (
    echo [i] %ISS_FILE% niet gevonden - overgeslagen
)
echo.
timeout /t 1 /nobreak >nul

REM ===== STAP 6: PYINSTALLER BUILD =====
cls
echo.
echo  ============================================
echo   STAP 6: PyInstaller Build
echo  ============================================
echo.

echo Opruimen van vorige build...
if exist "build"       rmdir /s /q "build"       2>nul
if exist "dist"        rmdir /s /q "dist"        2>nul
if exist "__pycache__" rmdir /s /q "__pycache__" 2>nul
echo [OK] Opgeruimd
echo.

echo Building...
pyinstaller "NVictEtiketten.spec"

if errorlevel 1 (
    echo.
    echo [!] PyInstaller build MISLUKT!
    pause & exit /b 1
)

set EXE_PATH=dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe
if not exist "%EXE_PATH%" (
    echo [!] Executable niet gevonden na build!
    pause & exit /b 1
)

for %%F in ("%EXE_PATH%") do set SIZE=%%~zF
set /a SIZE_MB=%SIZE% / 1048576
echo.
echo [OK] Build voltooid - EXE: %SIZE_MB% MB
echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 7: INNO SETUP =====
cls
echo.
echo  ============================================
echo   STAP 7: Inno Setup Installer
echo  ============================================
echo.

set "SETUP_FILE="
set "SETUP_NAME=NVict_Etiketten_Maker_Setup.exe"

if not defined ISCC (
    echo [i] Inno Setup niet gevonden - overgeslagen
    goto :skip_inno
)

if not exist "%ISS_FILE%" (
    echo [i] %ISS_FILE% niet gevonden - overgeslagen
    goto :skip_inno
)

if not exist "Output" mkdir Output

"%ISCC%" "%ISS_FILE%"

if errorlevel 1 (
    echo [!] Inno Setup MISLUKT - installer overgeslagen
    goto :skip_inno
)

if exist "Output\%SETUP_NAME%" (
    set "SETUP_FILE=Output\%SETUP_NAME%"
    for %%F in ("Output\%SETUP_NAME%") do set SIZE=%%~zF
    set /a SIZE_MB=!SIZE! / 1048576
    echo [OK] Installer: Output\%SETUP_NAME% - !SIZE_MB! MB
) else (
    echo [i] Installer bestand niet gevonden na build
)

:skip_inno
echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 8: VERSION JSON =====
cls
echo.
echo  ============================================
echo   STAP 8: Version JSON aanmaken
echo  ============================================
echo.

if not exist "Output" mkdir Output

if exist "Create_Version_JSON.ps1" (
    powershell -ExecutionPolicy Bypass -File "Create_Version_JSON.ps1" -Version "%VERSION%" -ReleaseNotesFile "release_notes.txt"
) else (
    echo [i] Create_Version_JSON.ps1 niet gevonden - handmatig aanmaken...
    (
        echo {
        echo   "version": "%VERSION%",
        echo   "download_url": "https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe",
        echo   "release_notes": "Versie %VERSION%",
        echo   "release_date": "%DATE%",
        echo   "update_check_url": "https://www.nvict.nl/software/updates/etiketten_version.json"
        echo }
    ) > Output\etiketten_version.json
)

if exist "Output\etiketten_version.json" (
    echo [OK] etiketten_version.json aangemaakt
    echo.
    type Output\etiketten_version.json
) else (
    echo [!] JSON aanmaken MISLUKT
)

echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 9: FTP UPLOAD =====
cls
echo.
echo  ============================================
echo   STAP 9: FTP Upload
echo  ============================================
echo.

REM Bepaal upload bestanden
set UPLOAD_COUNT=0
if defined SETUP_FILE (
    if exist "%SETUP_FILE%" set /a UPLOAD_COUNT+=1
)
if exist "Output\etiketten_version.json" set /a UPLOAD_COUNT+=1

if %UPLOAD_COUNT%==0 (
    echo [i] Geen bestanden om te uploaden - overgeslagen
    goto :skip_upload
)

echo Te uploaden bestanden:
if defined SETUP_FILE (
    if exist "%SETUP_FILE%" echo   [x] %SETUP_NAME%
)
if exist "Output\etiketten_version.json" echo   [x] etiketten_version.json
echo.

REM Lees FTP config
set "FTP_SERVER=ftp.nvict.nl"
set "FTP_USER=softwareupload@nvict.nl"
set "FTP_PATH=/NVict_Etiketten"
set "FTP_UPDATES_PATH=/software/updates"
set "FTP_PASS="

if exist "ftp_config.ini" (
    echo [i] ftp_config.ini geladen
    for /f "usebackq tokens=1,* delims==" %%a in ("ftp_config.ini") do (
        if "%%a"=="server"       set "FTP_SERVER=%%b"
        if "%%a"=="username"     set "FTP_USER=%%b"
        if "%%a"=="remote_path"  set "FTP_PATH=%%b"
        if "%%a"=="updates_path" set "FTP_UPDATES_PATH=%%b"
        if "%%a"=="password"     set "FTP_PASS=%%b"
    )
)

if not defined FTP_PASS (
    echo [i] Geen FTP wachtwoord in ftp_config.ini
    set /p "FTP_PASS=FTP wachtwoord: "
)
if "%FTP_PASS%"=="" (
    echo [i] Geen wachtwoord ingevoerd - upload overgeslagen
    goto :skip_upload
)

echo.
echo Server:  ftp://%FTP_SERVER%
echo Gebruiker: %FTP_USER%
echo.

REM Upload via WinSCP (primair) of PowerShell (fallback)
if defined WINSCP (
    echo [i] Uploaden via WinSCP...

    set "UPLOAD_SCRIPT=upload_etiketten_%RANDOM%.tmp"
    (
        echo option batch abort
        echo option confirm off
        echo open ftp://%FTP_USER%:%FTP_PASS%@%FTP_SERVER%
        if defined SETUP_FILE (
            echo cd %FTP_PATH%
            echo put "%SETUP_FILE%" "%SETUP_NAME%"
        )
        if exist "Output\etiketten_version.json" (
            echo cd %FTP_UPDATES_PATH%
            echo put "Output\etiketten_version.json"
        )
        echo close
        echo exit
    ) > "!UPLOAD_SCRIPT!"

    "%WINSCP%" /script="!UPLOAD_SCRIPT!"
    set UPLOAD_EXIT=!ERRORLEVEL!
    del "!UPLOAD_SCRIPT!" 2>nul

    if !UPLOAD_EXIT! NEQ 0 (
        echo [!] WinSCP upload MISLUKT - probeer PowerShell fallback...
        goto :upload_powershell
    )
    echo [OK] Upload via WinSCP voltooid
    set "UPLOAD_SUCCESS=J"
    goto :skip_upload
)

:upload_powershell
echo [i] Uploaden via PowerShell...

if exist "Upload_FTP_Etiketten.ps1" (
    powershell -ExecutionPolicy Bypass -File "Upload_FTP_Etiketten.ps1" ^
        -FtpServer "%FTP_SERVER%" ^
        -FtpUser "%FTP_USER%" ^
        -FtpPass "%FTP_PASS%" ^
        -FtpPath "%FTP_PATH%" ^
        -FtpUpdatesPath "%FTP_UPDATES_PATH%" ^
        -SetupFile "%SETUP_FILE%" ^
        -SetupName "%SETUP_NAME%"

    if !ERRORLEVEL! EQU 0 (
        echo [OK] Upload via PowerShell voltooid
        set "UPLOAD_SUCCESS=J"
    ) else (
        echo [!] PowerShell upload MISLUKT
    )
) else (
    echo [i] Upload_FTP_Etiketten.ps1 niet gevonden
    echo [i] Handmatig uploaden vereist
)

:skip_upload
echo.
timeout /t 2 /nobreak >nul

REM ===== SAMENVATTING =====
cls
echo.
echo  ============================================
echo   RELEASE VOLTOOID - v%VERSION%
echo  ============================================
echo.

if exist "dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe" (
    echo [OK] Applicatie:  dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe
)
if defined SETUP_FILE (
    if exist "%SETUP_FILE%" echo [OK] Installer:   %SETUP_FILE%
)
if exist "Output\etiketten_version.json" (
    echo [OK] Version JSON: Output\etiketten_version.json
)

if "%UPLOAD_SUCCESS%"=="J" (
    echo.
    echo [OK] FTP Upload geslaagd:
    if defined SETUP_FILE (
        echo      Setup:   https://www.nvict.nl/software/NVict_Etiketten/%SETUP_NAME%
    )
    echo      Version: https://www.nvict.nl/software/updates/etiketten_version.json
) else (
    echo.
    echo [i] FTP Upload niet uitgevoerd - handmatig uploaden nodig
)

echo.
echo  ============================================
echo.

if exist "Output" start "" "Output"

echo Release klaar!
echo.
pause
