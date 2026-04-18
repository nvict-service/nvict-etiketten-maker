# NVict Etiketten Maker

## Wat is het?

NVict Etiketten Maker is een Windows-applicatie waarmee je in een paar klikken
een volledig Excel-adressenbestand omzet naar drukklare adressstickers in Word.
Je kiest het formaat van je vel etiketten (Herma of Avery), selecteert je
Excel-bestand, en het programma genereert automatisch een keurig opgemaakt
Word-document waarin elk adres op de juiste plek op het vel staat — klaar om
te printen.

## Voor wie?

Voor iedereen die regelmatig post verstuurt naar een lijst mensen:
verenigingen, kleine bedrijven, ZZP'ers, kerkelijke organisaties,
fondsenwervers, of privégebruik rond feestdagen (kerstkaarten, uitnodigingen).
Geen technische kennis vereist — één knop, Excel erin, Word eruit.

## Wat maakt het bijzonder?

### Slimme auto-detectie

Je hoeft Excel-kolommen niet te hernoemen of in een specifieke volgorde te
zetten. Het programma herkent zelf welke kolom de achternaam, het adres, de
postcode en de woonplaats bevat — zowel Nederlandstalige als Engelstalige
kolomnamen (`aanhef`, `voorletters`, `straat`, `huisnummer`, `postcode`,
`woonplaats`, `land`, maar ook `surname`, `zip`, `city` etc.) — en plaatst ze
automatisch in de juiste volgorde op het etiket:

- Regel 1: aanhef, voorletters, achternaam
- Regel 2: eventuele tussenvoegsels / extra
- Regel 3: adres en huisnummer
- Regel 4: postcode en woonplaats
- Regel 5: land (alleen als die kolom bestaat)

### Preview vóór printen

Na het inladen zie je direct een schaal 1:1 voorbeeld van het eerste adres op
het gekozen etiket-formaat, zodat je precies weet hoe het eruit komt te zien
voordat er één vel papier wordt gebruikt.

### Zes standaardformaten vooraf ingesteld

| Formaat      | Afmeting         | Etiketten per vel |
|--------------|------------------|-------------------|
| Herma 10825  | 99,1 × 33,8 mm   | 16                |
| Herma 4625   | 66,0 × 33,8 mm   | 24                |
| Herma 4360   | 70,0 × 36,0 mm   | 24                |
| Herma 4267   | 99,1 × 42,3 mm   | 12                |
| Herma 4425   | 105,0 × 148,5 mm | 4                 |
| Avery L7163  | 99,1 × 38,1 mm   | 14                |

### Twee werkwijzen

1. **Excel-lijst** — elk adres op z'n eigen etiket, ideaal voor een mailing.
2. **Eén adres, heel vel** — handig voor retourstickers, afzendstickers of
   een vast bedrijfsadres.

### Modern uiterlijk

De app volgt automatisch je Windows-thema (licht/donker) en is ook handmatig
om te schakelen met één klik. De interface gebruikt moderne Segoe UI
typografie en een strakke, kaart-gebaseerde layout.

### Altijd up-to-date

Bij het opstarten checkt de app in de achtergrond of er een nieuwe versie
beschikbaar is. Zo ja, dan verschijnt er een subtiele notificatie in de
footer en kan de update met één klik gedownload en geïnstalleerd worden —
geen handmatig gedoe met downloads van de website.

## Wat krijg je als resultaat?

Een `.docx`-bestand (Microsoft Word) dat direct vanuit Word geprint kan
worden op je etikettenvel. Omdat het een Word-document is, kun je desgewenst
nog zelf kleine aanpassingen doen (een individueel etiket bewerken,
lettertype wijzigen, etc.) voordat je print.

## Vereisten

- Windows 10 of Windows 11
- Microsoft Word (om het resultaat te openen en te printen) — of een andere
  .docx-lezer
- Excel-bestand (`.xlsx` of `.xls`) als input; geen vaste kolomnamen vereist

## Installeren

De laatste getekende installer staat op:
<https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe>

De app wordt gesigneerd met een geldig codesigning-certificaat, dus Windows
SmartScreen geeft geen rode waarschuwing.

---

## Bouwen & uitrollen (voor ontwikkelaars)

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

### Versie

De actuele versie staat in `version.py` (`APP_VERSION`). De inhoud van
`release_notes.txt` wordt meegenomen in het auto-update manifest dat naar de
FTP-server wordt geupload.
