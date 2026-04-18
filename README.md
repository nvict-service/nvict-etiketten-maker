# NVict Etiketten Maker

Windows-applicatie voor het maken van adresstickers / etiketten.

## Bouwen & uitrollen

Deze app gebruikt de gedeelde build-toolkit in `../_nvict_build/`. Zie die
map voor de complete flow (PyInstaller → CodeSignTool → Inno Setup → FTPS).

```powershell
cd ..\_nvict_build
python release.py NVictEtikettenMaker               # volledige release
python release.py NVictEtikettenMaker --no-upload   # alleen lokaal builden
python release_ui.py                                # GUI met checkboxes
```

Of vanuit deze map:

```powershell
python build.py
```

## Versie

De actuele versie staat in `version.py` (`APP_VERSION`). Inhoud van
`release_notes.txt` wordt meegenomen in het auto-update manifest.

## Handmatig installeren

De laatste getekende installer staat op:
https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe
