@echo off
SETLOCAL ENABLEDELAYEDEXPANSION
REM ===================================================
REM FTP Upload Debug Script
REM ===================================================

color 0B
title FTP Upload Debug - NVict Etiketten

echo.
echo ================================================
echo FTP Upload Debug Script
echo ================================================
echo.

REM ===== STAP 1: Check WinSCP =====
echo [STAP 1] Zoeken naar WinSCP
echo ================================================
echo.

set "WINSCP="

if exist "C:\Program Files (x86)\WinSCP\WinSCP.com" (
    set "WINSCP=C:\Program Files (x86)\WinSCP\WinSCP.com"
    echo [OK] WinSCP gevonden (x86)
    echo      %WINSCP%
)

if exist "C:\Program Files\WinSCP\WinSCP.com" (
    set "WINSCP=C:\Program Files\WinSCP\WinSCP.com"
    echo [OK] WinSCP gevonden (x64)
    echo      %WINSCP%
)

if not defined WINSCP (
    echo [!] WinSCP NIET gevonden!
    echo.
    echo Gezocht in:
    echo   C:\Program Files (x86)\WinSCP\WinSCP.com
    echo   C:\Program Files\WinSCP\WinSCP.com
    echo.
    echo Oplossing:
    echo   Download WinSCP: https://winscp.net/eng/download.php
    echo   Installeer in standaard locatie
    echo.
    pause
    exit /b 1
)

echo.
pause

REM ===== STAP 2: Check ftp_config.ini =====
echo.
echo [STAP 2] Controleren ftp_config.ini
echo ================================================
echo.

if not exist "ftp_config.ini" (
    echo [!] ftp_config.ini NIET gevonden!
    echo.
    echo Oplossing:
    echo   1. Kopieer ftp_config.ini.template naar ftp_config.ini
    echo   2. Bewerk ftp_config.ini en vul je FTP wachtwoord in
    echo.
    echo Commando:
    echo   copy ftp_config.ini.template ftp_config.ini
    echo   notepad ftp_config.ini
    echo.
    pause
    exit /b 1
)

echo [OK] ftp_config.ini gevonden
echo.

REM Lees config
set "FTP_SERVER="
set "FTP_USER="
set "FTP_PASS="
set "FTP_PATH="
set "FTP_UPDATES_PATH="

echo Inhoud van ftp_config.ini:
echo ----------------------------------------
type ftp_config.ini
echo ----------------------------------------
echo.

for /f "usebackq tokens=1,* delims==" %%a in ("ftp_config.ini") do (
    if "%%a"=="server" set "FTP_SERVER=%%b"
    if "%%a"=="username" set "FTP_USER=%%b"
    if "%%a"=="remote_path" set "FTP_PATH=%%b"
    if "%%a"=="updates_path" set "FTP_UPDATES_PATH=%%b"
    if "%%a"=="password" set "FTP_PASS=%%b"
)

echo Gelezen configuratie:
echo   Server:       [%FTP_SERVER%]
echo   Username:     [%FTP_USER%]
echo   Remote path:  [%FTP_PATH%]
echo   Updates path: [%FTP_UPDATES_PATH%]
echo   Password:     [%FTP_PASS%]
echo.

if "%FTP_SERVER%"=="" (
    echo [!] Server niet ingesteld!
    pause
    exit /b 1
)

if "%FTP_USER%"=="" (
    echo [!] Username niet ingesteld!
    pause
    exit /b 1
)

if "%FTP_PASS%"=="" (
    echo [!] Password LEEG!
    echo.
    echo Je moet het wachtwoord invullen in ftp_config.ini
    echo.
    echo Bewerk het bestand:
    echo   notepad ftp_config.ini
    echo.
    echo En vul in bij password=:
    echo   password=JouwWachtwoordHier
    echo.
    pause
    exit /b 1
)

if "%FTP_PASS%"=="YOUR_PASSWORD_HERE" (
    echo [!] Password niet aangepast!
    echo.
    echo Je moet YOUR_PASSWORD_HERE vervangen door je echte wachtwoord
    echo.
    echo Bewerk het bestand:
    echo   notepad ftp_config.ini
    echo.
    pause
    exit /b 1
)

echo [OK] Configuratie lijkt compleet
echo.
pause

REM ===== STAP 3: Check bestanden om te uploaden =====
echo.
echo [STAP 3] Controleren bestanden
echo ================================================
echo.

set UPLOAD_COUNT=0

if exist "Output\NVict_Etiketten_Maker_Setup.exe" (
    echo [OK] Setup bestand gevonden
    for %%F in ("Output\NVict_Etiketten_Maker_Setup.exe") do (
        echo      Grootte: %%~zF bytes
        set SIZE=%%~zF
    )
    set /a SIZE_MB=!SIZE! / 1048576
    echo      Grootte: !SIZE_MB! MB
    set /a UPLOAD_COUNT+=1
) else (
    echo [!] Setup bestand NIET gevonden!
    echo      Verwacht: Output\NVict_Etiketten_Maker_Setup.exe
)

echo.

if exist "Output\etiketten_version.json" (
    echo [OK] Version JSON gevonden
    echo      Inhoud:
    type "Output\etiketten_version.json"
    set /a UPLOAD_COUNT+=1
) else (
    echo [!] Version JSON NIET gevonden!
    echo      Verwacht: Output\etiketten_version.json
)

echo.

if %UPLOAD_COUNT%==0 (
    echo [!] GEEN bestanden om te uploaden!
    echo.
    echo Run eerst het release script om bestanden te maken:
    echo   Release_Complete_Etiketten_v2.bat
    echo.
    pause
    exit /b 1
)

echo [OK] %UPLOAD_COUNT% bestand(en) klaar voor upload
echo.
pause

REM ===== STAP 4: Test FTP verbinding =====
echo.
echo [STAP 4] Testen FTP verbinding
echo ================================================
echo.

echo [i] Maken van test script...
set "TEST_SCRIPT=ftp_test_%RANDOM%.txt"

(
    echo option batch abort
    echo option confirm off
    echo open ftp://%FTP_USER%:%FTP_PASS%@%FTP_SERVER%
    echo ls
    echo close
    echo exit
) > "%TEST_SCRIPT%"

echo [i] Proberen te verbinden met FTP...
echo.

"%WINSCP%" /script="%TEST_SCRIPT%" /log="ftp_test.log"

set RESULT=%ERRORLEVEL%

if exist "ftp_test.log" (
    echo.
    echo FTP Log:
    echo ----------------------------------------
    type "ftp_test.log"
    echo ----------------------------------------
)

del "%TEST_SCRIPT%" 2>nul

echo.

if %RESULT% NEQ 0 (
    echo [!] FTP verbinding MISLUKT! (Error code: %RESULT%)
    echo.
    echo Mogelijke oorzaken:
    echo   - Verkeerd wachtwoord
    echo   - Verkeerde server/username
    echo   - FTP server niet bereikbaar
    echo   - Firewall blokkeert verbinding
    echo.
    echo Check ftp_test.log voor details
    echo.
    pause
    exit /b 1
)

echo [OK] FTP verbinding succesvol!
echo.
pause

REM ===== STAP 5: Probeer upload =====
echo.
echo [STAP 5] Uploaden naar FTP
echo ================================================
echo.

set "UPLOAD_SCRIPT=ftp_upload_%RANDOM%.txt"

(
    echo option batch abort
    echo option confirm off
    echo open ftp://%FTP_USER%:%FTP_PASS%@%FTP_SERVER%
    
    if exist "Output\NVict_Etiketten_Maker_Setup.exe" (
        echo cd %FTP_PATH%
        echo put "Output\NVict_Etiketten_Maker_Setup.exe"
    )
    
    if exist "Output\etiketten_version.json" (
        echo cd %FTP_UPDATES_PATH%
        echo put "Output\etiketten_version.json"
    )
    
    echo close
    echo exit
) > "%UPLOAD_SCRIPT%"

echo [i] Upload script:
type "%UPLOAD_SCRIPT%"
echo.
echo [i] Uploaden...
echo.

"%WINSCP%" /script="%UPLOAD_SCRIPT%" /log="ftp_upload.log"

set RESULT=%ERRORLEVEL%

if exist "ftp_upload.log" (
    echo.
    echo Upload Log:
    echo ----------------------------------------
    type "ftp_upload.log"
    echo ----------------------------------------
)

del "%UPLOAD_SCRIPT%" 2>nul

echo.

if %RESULT% NEQ 0 (
    echo [!] Upload MISLUKT! (Error code: %RESULT%)
    echo.
    echo Check ftp_upload.log voor details
    echo.
    pause
    exit /b 1
)

echo [OK] Upload succesvol!
echo.

REM ===== SAMENVATTING =====
echo.
echo ================================================
echo VOLTOOIING
echo ================================================
echo.

if exist "Output\NVict_Etiketten_Maker_Setup.exe" (
    echo [OK] Setup uploaded naar:
    echo      https://www.%FTP_SERVER%%FTP_PATH%/NVict_Etiketten_Maker_Setup.exe
)

if exist "Output\etiketten_version.json" (
    echo [OK] Version uploaded naar:
    echo      https://www.%FTP_SERVER%%FTP_UPDATES_PATH%/etiketten_version.json
)

echo.
echo Test de links in je browser!
echo.
pause
