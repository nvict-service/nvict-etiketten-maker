@echo off
SETLOCAL ENABLEDELAYEDEXPANSION
REM ===================================================
REM NVict Etiketten Maker - Complete Release v1.2
REM Met automatische release_notes.txt creatie
REM ===================================================

echo.
echo ========================================
echo NVict Etiketten Maker - Release v1.2
echo ========================================
echo.

REM ===== PYTHON FILE DETECTIE =====
echo [i] Zoeken naar Python bestand...
set PYTHON_FILE=

if exist "NVict_Etiketten.py" (
    set PYTHON_FILE=NVict_Etiketten.py
    echo [OK] Gevonden: NVict_Etiketten.py
)
if exist "NVict_Etiketten_v4.1.py" (
    if not defined PYTHON_FILE (
        set PYTHON_FILE=NVict_Etiketten_v4.1.py
        echo [OK] Gevonden: NVict_Etiketten_v4.1.py
        echo [!] TIP: Hernoem naar NVict_Etiketten.py voor productie
    )
)

if "%PYTHON_FILE%"=="" (
    echo [!] FOUT: Geen Python bestand gevonden!
    echo.
    pause
    exit /b 1
)

echo.

REM ===== VERSIE DETECTIE =====
echo [i] Versienummer detecteren uit %PYTHON_FILE%...
set VERSION=

for /f "tokens=2 delims== " %%a in ('findstr /C:"APP_VERSION = " "%PYTHON_FILE%"') do (
    set VERSION=%%a
)

set VERSION=%VERSION:"=%

if "%VERSION%"=="" (
    echo [!] Kon versienummer niet detecteren!
    pause
    exit /b 1
)

echo [OK] Versie gedetecteerd: %VERSION%
echo.

REM ===== RELEASE NOTES SETUP =====
echo ========================================
echo Release Notes Setup
echo ========================================
echo.

REM Check of release_notes.txt bestaat
if not exist "release_notes.txt" (
    echo [i] release_notes.txt niet gevonden
    echo [i] Aanmaken van template bestand...
    echo.
    
    REM Maak template release_notes.txt
    (
        echo • Automatische update controle bij opstarten
        echo • Download en installeer updates direct vanuit de app
        echo • Modern thema systeem ^(donker/licht mode^)
        echo • Verbeterde interface en dialogen
        echo • Snellere laadtijd en betere prestaties
        echo • Diverse bug fixes en verbeteringen
    ) > release_notes.txt
    
    echo [OK] release_notes.txt aangemaakt met template
    echo.
    echo Wilt u de release notes nu bewerken?
    echo.
    set /p "EDIT_NOTES=Open Notepad om te bewerken? (J/N): "
    
    if /i "!EDIT_NOTES!"=="J" (
        start /wait notepad.exe release_notes.txt
        echo [OK] Bestand opgeslagen
    )
    echo.
) else (
    echo [OK] Gevonden: release_notes.txt
    echo.
    echo Huidige inhoud:
    echo ----------------------------------------
    type release_notes.txt
    echo ----------------------------------------
    echo.
    echo Wilt u de release notes bewerken?
    echo.
    set /p "EDIT_NOTES=Open Notepad om te bewerken? (J/N): "
    
    if /i "!EDIT_NOTES!"=="J" (
        start /wait notepad.exe release_notes.txt
        echo [OK] Bestand opgeslagen
    )
    echo.
)

REM Lees release notes voor gebruik
set "NOTES="
for /f "usebackq delims=" %%a in ("release_notes.txt") do (
    if defined NOTES (
        set "NOTES=!NOTES! %%a"
    ) else (
        set "NOTES=%%a"
    )
)

if "!NOTES!"=="" set "NOTES=Versie %VERSION% release"

echo [OK] Release notes geladen
echo.
timeout /t 2 /nobreak >nul

REM ===== CONFIGURATIE =====
set SPEC_FILE=NVictEtiketten.spec
set ISS_FILE=NVict_Etiketten.iss

echo Configuratie:
echo   Versie:    %VERSION%
echo   Python:    %PYTHON_FILE%
echo   Spec:      %SPEC_FILE%
echo   ISS:       %ISS_FILE%
echo   Notes:     release_notes.txt
echo.
echo ========================================
echo   Het proces verloopt nu automatisch
echo ========================================
echo.
timeout /t 3 /nobreak >nul

REM ===== Pre-flight Checks =====
echo ========================================
echo Pre-flight Checks
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python NIET gevonden
    pause
    exit /b 1
)
echo [OK] Python

pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [!] PyInstaller NIET gevonden
    pause
    exit /b 1
)
echo [OK] PyInstaller

REM Inno Setup
set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    echo [OK] Inno Setup 6 (x86)
)
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
    echo [OK] Inno Setup 6 (x64)
)
if not defined ISCC (
    echo [i] Inno Setup niet gevonden (optioneel)
)

REM WinSCP
set "WINSCP="
if exist "C:\Program Files (x86)\WinSCP\WinSCP.com" (
    set "WINSCP=C:\Program Files (x86)\WinSCP\WinSCP.com"
    echo [OK] WinSCP (x86)
)
if exist "C:\Program Files\WinSCP\WinSCP.com" (
    set "WINSCP=C:\Program Files\WinSCP\WinSCP.com"
    echo [OK] WinSCP (x64)
)
if not defined WINSCP (
    echo [i] WinSCP niet gevonden (optioneel)
)

echo.
echo Benodigde bestanden:
if not exist "%SPEC_FILE%" (
    echo [!] %SPEC_FILE% NIET gevonden!
    pause
    exit /b 1
)
echo [OK] %SPEC_FILE%

if not exist "%ISS_FILE%" (
    echo [!] %ISS_FILE% NIET gevonden!
    pause
    exit /b 1
)
echo [OK] %ISS_FILE%

echo.
echo Alle checks voltooid!
echo.
timeout /t 2 /nobreak >nul

REM ===== VERSIE UPDATE IN ISS =====
echo.
echo ========================================
echo Versie Update in ISS bestand
echo ========================================
echo.

if exist "%ISS_FILE%" (
    echo [i] Update versie in %ISS_FILE% naar %VERSION%...
    
    copy "%ISS_FILE%" "%ISS_FILE%.backup" >nul 2>&1
    
    powershell -Command "(Get-Content '%ISS_FILE%') -replace '#define AppVersion \".*\"', '#define AppVersion \"%VERSION%\"' | Set-Content '%ISS_FILE%.tmp'"
    move /y "%ISS_FILE%.tmp" "%ISS_FILE%" >nul 2>&1
    
    echo [OK] Versie in ISS bestand bijgewerkt
)

echo.
timeout /t 1 /nobreak >nul

REM ===== STAP 1: PyInstaller =====
cls
echo.
echo ========================================
echo STAP 1: PyInstaller Build
echo ========================================
echo.

echo Opruimen oude bestanden...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "__pycache__" rmdir /s /q "__pycache__" 2>nul
echo [OK] Opgeruimd

echo.
echo Building met PyInstaller...
echo.
pyinstaller "%SPEC_FILE%"

if errorlevel 1 (
    echo.
    echo [!] Build MISLUKT!
    pause
    exit /b 1
)

if not exist "dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe" (
    echo.
    echo [!] NVict Etiketten Maker.exe niet gevonden!
    pause
    exit /b 1
)

echo.
echo [OK] Build voltooid!

for %%F in ("dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe") do (
    set SIZE=%%~zF
)
set /a SIZE_MB=%SIZE% / 1048576
echo      EXE grootte: %SIZE_MB% MB

echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 2: Inno Setup =====
cls
echo.
echo ========================================
echo STAP 2: Inno Setup
echo ========================================
echo.

if not defined ISCC (
    echo [i] Inno Setup niet gevonden - overgeslagen
    echo.
    goto skip_inno
)

echo [AUTOMATISCH] Installer wordt gemaakt...
echo.

"%ISCC%" "%ISS_FILE%"

if errorlevel 1 (
    echo.
    echo [!] Inno Setup MISLUKT!
    goto skip_inno
)

echo.
echo [OK] Installer gemaakt!

:skip_inno

set "SETUP_FILE="
set "SETUP_NAME="

if exist "Output\NVict_Etiketten_Maker_Setup.exe" (
    set "SETUP_FILE=Output\NVict_Etiketten_Maker_Setup.exe"
    set "SETUP_NAME=NVict_Etiketten_Maker_Setup.exe"
    echo.
    echo [OK] Setup: %SETUP_NAME%
    
    for %%F in ("Output\NVict_Etiketten_Maker_Setup.exe") do (
        set SIZE=%%~zF
    )
    set /a SIZE_MB=%SIZE% / 1048576
    echo      Grootte: %SIZE_MB% MB
)

echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 3: Version JSON =====
cls
echo.
echo ========================================
echo STAP 3: Version JSON
echo ========================================
echo.

if not exist "Output" mkdir Output

echo [AUTOMATISCH] Version JSON maken...
echo   Versie: %VERSION%
echo.

if not exist "Create_Version_JSON.ps1" (
    echo [!] Create_Version_JSON.ps1 niet gevonden
    echo [!] Maak JSON handmatig...
    
    (
        echo {
        echo   "version": "%VERSION%",
        echo   "download_url": "https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe",
        echo   "release_notes": "%NOTES%",
        echo   "release_date": "%DATE%",
        echo   "update_check_url": "https://www.nvict.nl/software/updates/etiketten_version.json"
        echo }
    ) > Output\etiketten_version.json
) else (
    powershell -ExecutionPolicy Bypass -File "Create_Version_JSON.ps1" -Version "%VERSION%" -ReleaseNotesFile "release_notes.txt"
)

echo.
echo [OK] Version JSON gemaakt
echo.
type Output\etiketten_version.json

echo.
timeout /t 2 /nobreak >nul

REM ===== STAP 4: FTP Upload =====
cls
echo.
echo ========================================
echo STAP 4: FTP Upload
echo ========================================
echo.

if not defined WINSCP (
    echo [i] WinSCP niet gevonden
    echo [i] Proberen via PowerShell FTP...
    echo.

    if exist "Upload_FTP.ps1" (
        powershell -ExecutionPolicy Bypass -File "Upload_FTP.ps1"
        goto skip_upload
    ) else (
        echo [i] Upload_FTP.ps1 ook niet gevonden - upload overgeslagen
        goto skip_upload
    )
)

set "FTP_SERVER=ftp.nvict.nl"
set "FTP_USER=softwareupload@nvict.nl"
set "FTP_PATH=/NVict_Etiketten"
set "FTP_UPDATES_PATH=/software/updates"
set "FTP_PASS="

if exist "ftp_config.ini" (
    echo [i] Config laden uit ftp_config.ini...
    
    for /f "usebackq tokens=1,* delims==" %%a in ("ftp_config.ini") do (
        if "%%a"=="server" set "FTP_SERVER=%%b"
        if "%%a"=="username" set "FTP_USER=%%b"
        if "%%a"=="remote_path" set "FTP_PATH=%%b"
        if "%%a"=="updates_path" set "FTP_UPDATES_PATH=%%b"
        if "%%a"=="password" set "FTP_PASS=%%b"
    )
)

echo.
echo FTP Settings:
echo   Server:       %FTP_SERVER%
echo   User:         %FTP_USER%
echo   Setup path:   %FTP_PATH%
echo   Updates path: %FTP_UPDATES_PATH%
echo.

set UPLOAD_COUNT=0
if defined SETUP_FILE (
    if exist "%SETUP_FILE%" set /a UPLOAD_COUNT+=1
)
if exist "Output\etiketten_version.json" set /a UPLOAD_COUNT+=1

if %UPLOAD_COUNT%==0 (
    echo [i] Geen bestanden om te uploaden
    goto skip_upload
)

echo Te uploaden:
if defined SETUP_FILE (
    if exist "%SETUP_FILE%" echo   [x] %SETUP_NAME% -^> %FTP_PATH%
)
if exist "Output\etiketten_version.json" echo   [x] etiketten_version.json -^> %FTP_UPDATES_PATH%
echo.

if not defined FTP_PASS (
    echo [i] Geen FTP wachtwoord in ftp_config.ini
    echo [i] Upload overgeslagen
    goto skip_upload
)

if "%FTP_PASS%"=="" (
    echo [i] FTP wachtwoord leeg - upload overgeslagen
    goto skip_upload
)

echo [AUTOMATISCH] Uploaden naar FTP...
echo.

set "UPLOAD_SCRIPT=upload_%RANDOM%.txt"

(
    echo option batch abort
    echo option confirm off
    echo open ftp://%FTP_USER%:%FTP_PASS%@%FTP_SERVER%
    
    if defined SETUP_FILE (
        if exist "%SETUP_FILE%" (
            echo cd %FTP_PATH%
            echo put "%SETUP_FILE%" "%SETUP_NAME%"
        )
    )
    
    if exist "Output\etiketten_version.json" (
        echo cd %FTP_UPDATES_PATH%
        echo put "Output\etiketten_version.json"
    )
    
    echo close
    echo exit
) > "%UPLOAD_SCRIPT%"

"%WINSCP%" /script="%UPLOAD_SCRIPT%"

if errorlevel 1 (
    echo.
    echo [!] Upload MISLUKT!
    del "%UPLOAD_SCRIPT%" 2>nul
    goto skip_upload
)

del "%UPLOAD_SCRIPT%" 2>nul

echo.
echo [OK] Upload voltooid!
set "UPLOAD_SUCCESS=J"

:skip_upload

echo.
timeout /t 2 /nobreak >nul

REM ===== SAMENVATTING =====
cls
echo.
echo ========================================
echo VOLTOOIING
echo ========================================
echo.

if exist "dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe" (
    echo [OK] Applicatie: dist\NVict_Etiketten_Maker\NVict Etiketten Maker.exe
)

if defined SETUP_FILE (
    if exist "%SETUP_FILE%" (
        echo [OK] Installer:   %SETUP_FILE%
    )
)

if exist "Output\etiketten_version.json" (
    echo [OK] Version:     Output\etiketten_version.json
)

if "%UPLOAD_SUCCESS%"=="J" (
    echo [OK] Upload:      
    if defined SETUP_FILE (
        echo      Setup:   https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe
    )
    if exist "Output\etiketten_version.json" (
        echo      Version: https://www.nvict.nl/software/updates/etiketten_version.json
    )
)

echo.
echo ========================================
echo.

if exist "Output" (
    echo Opening Output directory...
    start "" "Output"
)

echo.
echo Release proces voltooid!
echo.
pause
