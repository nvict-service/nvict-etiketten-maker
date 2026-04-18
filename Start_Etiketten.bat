@echo off
REM ===================================================
REM NVict Etiketten Maker - Start Script
REM Controleert en installeert dependencies automatisch
REM ===================================================

echo.
echo =====================================
echo  NVict Etiketten Maker starten...
echo =====================================
echo.

REM Controleer Python
py --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo [!] Python is niet geinstalleerd!
        echo.
        echo Download Python via: https://www.python.org/downloads/
        echo Zorg dat "Add Python to PATH" aangevinkt is tijdens installatie.
        echo.
        pause
        exit /b 1
    )
    set PYTHON=python
) else (
    set PYTHON=py
)

echo [OK] Python gevonden:
%PYTHON% --version
echo.

REM Controleer en installeer modules
echo Benodigde modules controleren...
echo.

%PYTHON% -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo [ ] pandas niet gevonden - installeren...
    %PYTHON% -m pip install pandas --quiet
    if errorlevel 1 ( echo [!] pandas installatie mislukt! & goto :install_error )
    echo [OK] pandas geinstalleerd
) else (
    echo [OK] pandas
)

%PYTHON% -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo [ ] python-docx niet gevonden - installeren...
    %PYTHON% -m pip install python-docx --quiet
    if errorlevel 1 ( echo [!] python-docx installatie mislukt! & goto :install_error )
    echo [OK] python-docx geinstalleerd
) else (
    echo [OK] python-docx
)

%PYTHON% -c "from PIL import Image" >nul 2>&1
if errorlevel 1 (
    echo [ ] Pillow niet gevonden - installeren...
    %PYTHON% -m pip install Pillow --quiet
    if errorlevel 1 ( echo [!] Pillow installatie mislukt! & goto :install_error )
    echo [OK] Pillow geinstalleerd
) else (
    echo [OK] Pillow
)

%PYTHON% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [ ] openpyxl niet gevonden - installeren...
    %PYTHON% -m pip install openpyxl --quiet
    if errorlevel 1 ( echo [!] openpyxl installatie mislukt! & goto :install_error )
    echo [OK] openpyxl geinstalleerd
) else (
    echo [OK] openpyxl
)

echo.
echo Alle modules aanwezig. App starten...
echo.

REM Start de applicatie
%PYTHON% NVict_Etiketten.py

echo.
echo App afgesloten.
pause
exit /b 0

:install_error
echo.
echo =====================================
echo [!] Installatie MISLUKT
echo =====================================
echo.
echo Probeer handmatig:
echo   %PYTHON% -m pip install pandas python-docx Pillow openpyxl
echo.
echo Of zorg dat u een internetverbinding heeft.
echo.
pause
exit /b 1
