# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec bestand voor NVict Etiketten Maker
# Versie: 4.0.0 - Multi-format Edition met Voorbeeld Bestanden

import os

# BELANGRIJKE NOTITIE:
# De voorbeeld Excel bestanden (Voorbeeld_Adressen.xlsx, etc.) worden NIET 
# meegebundeld in de .exe door PyInstaller. Ze worden door de Inno Setup 
# installer geplaatst in: C:\Users\[User]\Documents\Etiketten Maker\
#
# Het programma zoekt automatisch naar deze map via get_voorbeelden_map()

# 1. ANALYSE FASE
# Hier definiëren we alle bestanden die nodig zijn.
a = Analysis(
    ['NVict_Etiketten.py'],
    pathex=[],
    binaries=[],
    # DATAS: Bundel alle bronbestanden mee in de .exe
    datas=[
        ('favicon.ico', '.'),     # Het hoofd applicatie icoon
        ('logo.png', '.'),        # Het logo voor de header (optioneel)
        
        # OPMERKING: Voorbeeld bestanden worden NIET hier toegevoegd!
        # Ze worden door Inno Setup installer geplaatst in Documenten map.
        # Voor lokaal testen kun je ze optioneel uncomment-en:
        # ('Voorbeeld_Adressen.xlsx', '.'),
        # ('Voorbeeld_Adressen_Uitgebreid.xlsx', '.'),
        # ('LEES_MIJ_VOORBEELDEN.txt', '.'),
    ],
    hiddenimports=[
        'winreg',           # Voor Windows thema detectie
        'pandas',           # Voor Excel verwerking
        'openpyxl',         # Voor .xlsx bestanden
        'xlrd',             # Voor .xls bestanden
        'docx',             # Voor Word document creatie
        'PIL',              # Voor logo afbeeldingen
        'urllib.request',   # Voor update check
        'urllib.error',     # Voor update check error handling
        'json',             # Voor update check JSON parsing
        'threading',        # Voor achtergrond update check
        'webbrowser',       # Voor openen van download links
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# 2. EXE FASE
# Dit creëert de uitvoerbare (EXE) file.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NVict Etiketten Maker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,                # Geen console venster
    
    # BELANGRIJK: Dit stelt het icoon voor het EXE-bestand in
    icon='favicon.ico', 
    
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# 3. COLLECT FASE
# Dit verzamelt alles in de uiteindelijke map.
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='NVict_Etiketten_Maker',
)
