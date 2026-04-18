; Script: NVict_Etiketten_Setup.iss
; Doel: Installeert NVict Etiketten Maker - Multi-formaat etiket generator
; Versie: 4.0.0

;
;--- PREPROCESSOR DEFINITIES ---
#define AppName "NVict Etiketten Maker"
#define AppVersion "4.0.0" 
#define AppPublisher "NVict Service"
#define AppURL "https://www.nvict.nl"
#define AppExeName "NVict Etiketten Maker.exe"
#define AppUninsKey "NVict_Etiketten_Maker_App" 

[Setup]
;
;--- Algemene Instellingen ---
AppName={#AppName} 
AppVersion={#AppVersion} 
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
AppId="{#AppUninsKey}"
DefaultDirName={autopf}\NVict\Etiketten
DefaultGroupName=NVict Service
OutputBaseFilename=NVict_Etiketten_Maker_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}  
VersionInfoVersion={#AppVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Installatieprogramma
VersionInfoCopyright=Copyright (C) 2024 {#AppPublisher}

; Icoon voor de installer zelf
SetupIconFile=C:\Users\NVict Service\OneDrive - NVict Service\Apps\NVictPython\NVict-Etiketten\favicon.ico 

; Wizard afbeeldingen (optioneel)
; WizardImageFile=compiler:WizModernImage-IS.bmp
; WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp

; Rechten
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "nl"; MessagesFile: "compiler:Languages\Dutch.isl"

;
;--- Bestanden en Mappen ---
[Files]
; Installeert alle bestanden uit de PyInstaller 'dist' map
; Voor one-folder versie:
Source: "C:\Users\NVict Service\OneDrive - NVict Service\Apps\NVictPython\NVict-Etiketten\dist\NVict_Etiketten_Maker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; OF voor one-file versie (kies één van de twee):
; Source: "C:\Users\NVict Service\OneDrive - NVict Service\Apps\NVictPython\NVict-Etiketten\dist\NVict Etiketten Maker.exe"; DestDir: "{app}"; Flags: ignoreversion

; Extra bestanden die nuttig kunnen zijn (optioneel)
; Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
; Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

; NIEUW: Voorbeeld Excel bestanden naar Documenten\Etiketten Maker
; Deze map wordt automatisch aangemaakt als hij nog niet bestaat
Source: "Voorbeeld_Adressen.xlsx"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist
Source: "Voorbeeld_Adressen_Uitgebreid.xlsx"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist
Source: "LEES_MIJ_VOORBEELDEN.txt"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist

; --- Snelkoppelingen ---
[Icons]
; Start Menu snelkoppeling
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Comment: "Excel naar etiketten converteren"

; Desktop snelkoppeling
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

; Snelkoppeling voor deinstallatie
Name: "{group}\Verwijder {#AppName}"; Filename: "{uninstallexe}"

; NIEUW: Snelkoppeling naar de voorbeelden map
Name: "{group}\Voorbeeld Bestanden"; Filename: "{userdocs}\Etiketten Maker"; Comment: "Open de map met voorbeeld Excel bestanden"

; --- Taken (Optionele installatie opties) ---
[Tasks]
Name: "desktopicon"; Description: "Maak een snelkoppeling op het bureaublad"; GroupDescription: "Extra iconen:"; Flags: unchecked

; --- Registry Entries ---
[Registry]
; Basisregistratie voor geïnstalleerde applicaties
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#AppVersion}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "ExePath"; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey

; Optioneel: URL Protocol Handler (voor toekomstig gebruik)
; Root: HKCR; Subkey: "nvict-etiketten"; ValueType: string; ValueName: ""; ValueData: "URL:NVict Etiketten Protocol"; Flags: uninsdeletekey
; Root: HKCR; Subkey: "nvict-etiketten"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""
; Root: HKCR; Subkey: "nvict-etiketten\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""

; --- Acties na Installatie ---
[Run]
; Optie om applicatie direct te starten na installatie
Filename: "{app}\{#AppExeName}"; Description: "Start {#AppName}"; Flags: nowait postinstall skipifsilent

; --- Code Sectie ---
[Code]
// Functie om te controleren of er al een versie geïnstalleerd is
function InitializeSetup(): Boolean;
var
  OldVersion: String;
  OldPath: String;
begin
  Result := True;
  
  // Check of er een oude versie is
  if RegQueryStringValue(HKLM, 'Software\{#AppPublisher}\{#AppName}', 'Version', OldVersion) then
  begin
    if MsgBox('Er is al een versie geïnstalleerd (' + OldVersion + ').' + #13#10 + 
              'Wilt u doorgaan met de installatie van versie {#AppVersion}?', 
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;

// Functie die wordt uitgevoerd na succesvolle installatie
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Hier kun je extra acties toevoegen na installatie
    // Bijvoorbeeld: configuratiebestanden aanmaken, registry keys, etc.
  end;
end;

// Functie die wordt uitgevoerd na deinstallatie
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Opruimen van gebruikersdata (optioneel)
    // Bijvoorbeeld: verwijderen van configuratiebestanden in AppData
    // DelTree(ExpandConstant('{userappdata}\NVict\Etiketten'), True, True, True);
  end;
end;

// Custom wizard pagina's (optioneel)
// Bijvoorbeeld voor licentie overeenkomst, readme, enz.
