# NVict Etiketten Maker - Professional Edition v4.1

## 📋 Overzicht

Complete deployment pipeline voor NVict Etiketten Maker met:
- ✅ Automatische versie detectie
- ✅ Geautomatiseerde build (PyInstaller)
- ✅ Installer creatie (Inno Setup)
- ✅ Version JSON generatie
- ✅ Automatische FTP upload
- ✅ Geavanceerd update systeem in de app

---

## 🚀 Quick Start

### 1. Voorbereiding (éénmalig)

#### Installeer benodigde tools:
```bash
# Python packages
pip install pyinstaller pandas python-docx pillow openpyxl

# Download en installeer:
- Inno Setup 6: https://jrsoftware.org/isdl.php
- WinSCP: https://winscp.net/eng/download.php (voor FTP upload)
```

#### Configureer FTP (optioneel):
1. Kopieer `ftp_config.ini.template` naar `ftp_config.ini`
2. Vul je FTP wachtwoord in
3. **BELANGRIJK**: Voeg `ftp_config.ini` toe aan `.gitignore`!

### 2. Release Maken

```bash
# Stap 1: Update versie nummer in NVict_Etiketten.py
# Pas APP_VERSION aan naar nieuwe versie (bijv. "4.1.0")

# Stap 2: Run het release script
Release_Complete_Etiketten.bat

# Stap 3: Volg de instructies
# - Kies hoe je release notes wilt invoeren
# - Daarna verloopt alles automatisch!
```

### 3. Resultaat

Na voltooiing vind je:
- `dist\NVict_Etiketten_Maker\` - Complete applicatie
- `Output\NVict_Etiketten_Maker_Setup.exe` - Installer
- `Output\etiketten_version.json` - Version info voor updates

---

## 📁 Project Structuur

```
NVict_Etiketten/
├── NVict_Etiketten.py              # Hoofdbestand met APP_VERSION
├── NVictEtiketten.spec             # PyInstaller configuratie
├── NVict_Etiketten.iss             # Inno Setup script
├── Release_Complete_Etiketten.bat  # Geautomatiseerd release script
├── Create_Version_JSON.ps1         # JSON generator voor updates
├── ftp_config.ini                  # FTP credentials (NIET committen!)
├── ftp_config.ini.template         # Template voor FTP config
├── favicon.ico                     # App icoon
├── Logo.png                        # Logo voor in app
├── release_notes.txt               # Release notes (optioneel)
│
├── dist/                           # PyInstaller output
├── build/                          # PyInstaller temp
└── Output/                         # Installer output
```

---

## ⚙️ Geavanceerde Features

### Update Systeem

De app heeft een compleet update systeem:

1. **Automatische Check bij Opstarten**
   - Checkt na 2 seconden of er een update is
   - Gebruikt fallback URL bij blokkering
   - Toont update notificatie in footer

2. **Update Dialog**
   - Toont huidige vs nieuwe versie
   - Geeft release notes weer
   - 3 opties: Download & Installeer / Alleen Download / Later

3. **Download & Install**
   - Download installer automatisch
   - Toont progress
   - Start installer na download
   - Sluit app automatisch af

### Version JSON Format

```json
{
  "version": "4.1.0",
  "download_url": "https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe",
  "release_notes": "• Multi-formaat ondersteuning\n• Geavanceerd update systeem\n• Verbeterde UI",
  "release_date": "2025-01-17",
  "update_check_url": "https://www.nvict.nl/software/updates/etiketten_version.json"
}
```

---

## 🔧 Release Script Details

### Release_Complete_Etiketten.bat

Het script voert automatisch de volgende stappen uit:

#### Stap 1: Versie Detectie
- Leest `APP_VERSION` uit Python bestand
- Geen handmatige invoer nodig!

#### Stap 2: Release Notes
3 opties voor release notes:
1. **Kort** - Type direct een regel
2. **Lang** - Notepad opent voor meerdere regels
3. **Bestand** - Lees uit `release_notes.txt`

#### Stap 3: Pre-flight Checks
Controleert:
- ✓ Python geïnstalleerd
- ✓ PyInstaller beschikbaar
- ✓ Inno Setup gevonden
- ✓ WinSCP aanwezig
- ✓ Alle benodigde bestanden

#### Stap 4: PyInstaller Build
- Ruimt oude builds op
- Build one-folder distributie
- Toont EXE grootte
- Controleert output

#### Stap 5: Inno Setup
- Update versie in ISS bestand
- Compileer installer
- Toont setup grootte

#### Stap 6: Version JSON
- Genereert JSON met PowerShell
- Gebruikt UTF-8 zonder BOM
- Leest release notes uit bestand/variabele

#### Stap 7: FTP Upload
- Upload setup naar `/NVict_Etiketten/`
- Upload JSON naar `/software/updates/`
- Toont progress en status

#### Stap 8: Samenvatting
- Lijst van aangemaakte bestanden
- URLs van geüploade bestanden
- Opent Output directory

---

## 🎨 UI Thema Systeem

De app heeft een modern thema systeem:

### Dark Mode
```python
bg_primary = "#202020"
bg_secondary = "#2d2d2d"
text_primary = "#ffffff"
accent = "#0078d4"
```

### Light Mode
```python
bg_primary = "#f3f3f3"
bg_secondary = "#ffffff"
text_primary = "#202020"
accent = "#0078d4"
```

### Automatische Detectie
- Leest Windows thema uit registry
- Gebruiker kan handmatig wisselen
- Alle UI elementen passen zich aan

---

## 🌐 FTP Upload

### Configuratie

Maak `ftp_config.ini`:
```ini
server=ftp.nvict.nl
username=softwareupload@nvict.nl
password=JOUW_WACHTWOORD
remote_path=/NVict_Etiketten
updates_path=/software/updates
```

### Upload Structuur

```
FTP Server:
├── /NVict_Etiketten/
│   └── NVict_Etiketten_Maker_Setup.exe    # Installer
│
└── /software/updates/
    └── etiketten_version.json              # Version info
```

### Security

⚠️ **BELANGRIJK**:
- Bewaar `ftp_config.ini` NOOIT in git!
- Voeg toe aan `.gitignore`
- Gebruik sterke wachtwoorden
- Test connectie eerst handmatig

---

## 📦 PyInstaller Details

### Spec File Opties

```python
console=False           # GUI app, geen console
upx=True               # Compressie met UPX
icon='favicon.ico'     # App icoon
```

### Excluded Modules
Voor kleinere executable:
```python
excludes=[
    'matplotlib',      # Niet nodig voor onze app
    'scipy',
    'jupyter',
]
```

### Hidden Imports
Zorg dat alles werkt:
```python
hiddenimports=[
    'PIL._tkinter_finder',
    'openpyxl',
    'python-docx',
]
```

---

## 🔨 Inno Setup Details

### App ID
```pascal
AppId={{E5A2B9F3-7D4C-4A8E-9B2F-1C8D6E5A4F3B}
```
Unieke GUID voor deze applicatie

### Installatie Locaties
- **App**: `C:\Program Files\NVict\Etiketten\`
- **Voorbeelden**: `%UserProfile%\Documents\Etiketten Maker\`
- **Start Menu**: `NVict Service`

### Registry Keys
```
HKLM\Software\NVict Service\NVict Etiketten Maker\
├── InstallPath
├── Version
└── ExePath
```

### Features
- ✅ Automatische upgrade (verwijdert oude versie)
- ✅ Desktop icoon (optioneel)
- ✅ Voorbeeld bestanden in Documenten
- ✅ Start Menu snelkoppelingen

---

## 🐛 Troubleshooting

### Build Problemen

**PyInstaller fout: "Module not found"**
```bash
pip install --upgrade <module>
pip install --upgrade pyinstaller
```

**UPX comprressie fout**
```python
# In .spec file:
upx=False  # Schakel UPX uit
```

### Installer Problemen

**Inno Setup kan bestanden niet vinden**
- Check of relatieve paden kloppen
- Controleer of `dist\NVict_Etiketten_Maker\` bestaat
- Run `Release_Complete_Etiketten.bat` opnieuw

**Setup werkt niet na installatie**
- Check of alle dependencies in spec file staan
- Test executable eerst in `dist\` folder

### FTP Problemen

**Upload mislukt: "Connection refused"**
- Controleer firewall instellingen
- Test connectie in FileZilla
- Verifieer FTP credentials

**Timeout tijdens upload**
- Check internet verbinding
- Grote bestanden kunnen lang duren
- Verhoog timeout in WinSCP

### Update Systeem Problemen

**Update wordt niet gedetecteerd**
- Check of JSON URL bereikbaar is
- Test JSON format op https://jsonlint.com
- Verifieer versie nummering (x.y.z format)

**Download mislukt**
- Controleer download URL in JSON
- Test URL handmatig in browser
- Check firewall/antivirus

---

## 📝 Versie Nummering

Gebruik Semantic Versioning:
- **MAJOR** (4.x.x) - Breaking changes
- **MINOR** (x.1.x) - Nieuwe features
- **PATCH** (x.x.0) - Bug fixes

Voorbeelden:
```
4.0.0 → 4.1.0  # Nieuwe features toegevoegd
4.1.0 → 4.1.1  # Bug fix
4.1.1 → 5.0.0  # Breaking change
```

---

## 🔐 Security Checklist

- [ ] `ftp_config.ini` in `.gitignore`
- [ ] Geen wachtwoorden in scripts
- [ ] Test downloads via HTTPS
- [ ] Verifieer installer signatures
- [ ] Check bestandspermissies op server
- [ ] Regular security updates

---

## 📞 Support

**Website**: www.nvict.nl  
**Email**: support@nvict.nl  

---

## 📜 License

Copyright © 2025 NVict Service  
All rights reserved.

---

## ✨ Credits

Ontwikkeld door NVict Service met liefde en Python 🐍
