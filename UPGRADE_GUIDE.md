# 🚀 Upgrade Guide: NVict Etiketten v4.0 → v4.1

## Wat is er nieuw in v4.1?

### ✨ Nieuwe Features
1. **Geavanceerd Update Systeem**
   - Automatische update check bij opstarten
   - Download & Install functionaliteit
   - Progress dialogs
   - Release notes weergave

2. **Complete Deployment Pipeline**
   - Automatische versie detectie
   - Geautomatiseerde build process
   - FTP upload functionaliteit
   - Version JSON generatie

3. **Verbeterde Code Structuur**
   - Consistent theme systeem
   - Better resource handling
   - Improved error handling
   - Professional dialog windows

---

## 📋 Upgrade Stappen

### Stap 1: Backup je huidige versie
```bash
# Kopieer je huidige werkmap
cp -r NVict-Etiketten NVict-Etiketten-BACKUP

# Of in Windows:
xcopy NVict-Etiketten NVict-Etiketten-BACKUP /E /I
```

### Stap 2: Download nieuwe bestanden

Vervang de volgende bestanden met de nieuwe versies:

#### Hoofdbestanden
- ✅ `NVict_Etiketten.py` → Nieuwe versie met update systeem
- ✅ `NVictEtiketten.spec` → Verbeterde PyInstaller config
- ✅ `NVict_Etiketten.iss` → Verbeterde Inno Setup script

#### Nieuwe deployment bestanden
- ⭐ `Release_Complete_Etiketten.bat` → Geautomatiseerd release script
- ⭐ `Create_Version_JSON.ps1` → Version JSON generator
- ⭐ `ftp_config.ini.template` → FTP configuratie template
- ⭐ `requirements.txt` → Python dependencies
- ⭐ `.gitignore` → Git ignore regels
- ⭐ `DEPLOYMENT_README.md` → Complete documentatie

### Stap 3: Update je environment

```bash
# Installeer/update dependencies
pip install -r requirements.txt

# Of handmatig:
pip install --upgrade pandas python-docx Pillow openpyxl pyinstaller
```

### Stap 4: Configureer FTP (optioneel)

Als je automatische uploads wilt:

```bash
# 1. Kopieer template
copy ftp_config.ini.template ftp_config.ini

# 2. Edit ftp_config.ini en vul je credentials in

# 3. Voeg toe aan .gitignore (als je git gebruikt)
echo ftp_config.ini >> .gitignore
```

### Stap 5: Test de nieuwe versie

```bash
# Run de app direct om te testen
python NVict_Etiketten.py
```

Controleer:
- ✅ App start correct
- ✅ Theme systeem werkt
- ✅ Excel import werkt
- ✅ Etiketten genereren werkt
- ✅ Update check werkt (na 2 seconden)

### Stap 6: Maak je eerste release

```bash
# 1. Update versie in NVict_Etiketten.py
# Verander APP_VERSION = "4.0.0" naar "4.1.0"

# 2. Run release script
Release_Complete_Etiketten.bat

# 3. Volg de wizard
# - Kies release notes optie
# - Alles gebeurt automatisch!
```

---

## 🔄 Migratie Checklist

### Voor je begint
- [ ] Backup maken van oude versie
- [ ] Python 3.8+ geïnstalleerd
- [ ] Inno Setup 6 geïnstalleerd
- [ ] WinSCP geïnstalleerd (voor FTP)
- [ ] Git configuratie up-to-date (optioneel)

### Tijdens migratie
- [ ] Alle nieuwe bestanden gedownload
- [ ] Dependencies geïnstalleerd
- [ ] FTP geconfigureerd (optioneel)
- [ ] App getest lokaal

### Na migratie
- [ ] Eerste release gemaakt
- [ ] Installer getest
- [ ] Update systeem getest
- [ ] Documentatie gelezen

---

## 📝 Belangrijke Verschillen

### Oude structuur (v4.0)
```
NVict-Etiketten/
├── NVict_Etiketten.py
├── NVict_Etiketten.iss
└── (handmatige build process)
```

### Nieuwe structuur (v4.1)
```
NVict-Etiketten/
├── NVict_Etiketten.py              # Met update systeem
├── NVictEtiketten.spec             # PyInstaller config
├── NVict_Etiketten.iss             # Verbeterd
├── Release_Complete_Etiketten.bat  # Geautomatiseerd!
├── Create_Version_JSON.ps1         # Nieuwe helper
├── ftp_config.ini                  # FTP credentials
├── requirements.txt                # Dependencies
├── .gitignore                      # Git ignore
└── DEPLOYMENT_README.md            # Documentatie
```

---

## 🎯 Veelgestelde Vragen

### Q: Moet ik alles overzetten?
**A:** Minimaal nodig:
- `NVict_Etiketten.py` (nieuwe versie)
- `NVictEtiketten.spec`
- `NVict_Etiketten.iss`
- `Release_Complete_Etiketten.bat`

De rest is optioneel maar sterk aanbevolen!

### Q: Werkt de oude .iss file nog?
**A:** Ja, maar de nieuwe heeft betere features:
- Automatische versie update
- Beter upgrade handling
- Relatieve paden
- Voorbeelden installatie

### Q: Moet ik FTP gebruiken?
**A:** Nee, dat is optioneel. Je kunt:
- Handmatig uploaden via FileZilla
- Andere upload methode gebruiken
- Helemaal geen upload doen

### Q: Kan ik terug naar v4.0?
**A:** Ja, gebruik je backup:
```bash
# Herstel backup
rm -rf NVict-Etiketten
mv NVict-Etiketten-BACKUP NVict-Etiketten
```

### Q: Wat als ik git gebruik?
**A:** Gebruik de nieuwe `.gitignore`:
```bash
# Voeg toe aan je repo
cp .gitignore /path/to/your/repo/

# Commit nieuwe bestanden
git add .
git commit -m "Upgrade naar v4.1"
```

---

## 🆘 Problemen?

### Build faalt
```bash
# Reinstall dependencies
pip uninstall -y pandas python-docx Pillow openpyxl
pip install -r requirements.txt

# Clean en probeer opnieuw
rmdir /s /q build dist
Release_Complete_Etiketten.bat
```

### App crasht
```bash
# Test zonder PyInstaller
python NVict_Etiketten.py

# Check error messages
# Als het lokaal werkt, rebuild:
pyinstaller --clean NVictEtiketten.spec
```

### FTP werkt niet
```bash
# Test credentials in FileZilla
# Check ftp_config.ini syntax
# Verifieer internet verbinding
```

---

## 📚 Meer Hulp Nodig?

Lees de volledige documentatie:
- `DEPLOYMENT_README.md` - Complete deployment guide
- Release script comments - Inline uitleg
- Inno Setup file - Pascal code comments

---

## ✅ Je bent klaar!

Na deze upgrade heb je:
- ✨ Moderne update systeem
- 🚀 Geautomatiseerde deployment
- 📦 Professional installer
- 🌐 FTP upload mogelijkheid
- 📝 Complete documentatie

**Volgende stappen:**
1. Maak je eerste release met het nieuwe systeem
2. Test het update mechanisme
3. Upload naar je website
4. Geniet van de automated workflow! 🎉

---

**Vragen of problemen?**  
📧 support@nvict.nl  
🌐 www.nvict.nl
