; NVict Etiketten Maker Installer Script met RELATIEVE paden
; Versie 4.2 — gemigreerd naar _nvict_build pipeline
;
; Wordt aangeroepen door _nvict_build/release.py met /DVERSION=... en /Ssigntool=...
; Handmatig compileren: ISCC NVict_Etiketten.iss /DVERSION=4.2.0

#ifndef VERSION
  #define VERSION "4.2.0"
#endif

#define AppName "NVict Etiketten Maker"
#define AppVersion VERSION
#define AppPublisher "NVict Service"
#define AppURL "https://www.nvict.nl"
#define AppExeName "NVict Etiketten Maker.exe"
#define AppUninsKey "NVict_Etiketten_Maker_App"

[Setup]
; --- SIGNING ---
; De setup.exe wordt NA Inno gesigned door release.py (CodeSignTool).
; De geembedde uninstaller wordt niet gesigned — zelfde gedrag als het
; oude Release_Complete_v3_0.bat script.

; APP INFORMATIE
AppId={{E5A2B9F3-7D4C-4A8E-9B2F-1C8D6E5A4F3B}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}

; INSTALLATIE DIRECTORIES
DefaultDirName={autopf}\NVict\Etiketten
DefaultGroupName=NVict Service
DisableProgramGroupPage=yes

; OUTPUT INSTELLINGEN
OutputDir=Output
OutputBaseFilename=NVict_Etiketten_Maker_Setup
Compression=lzma
SolidCompression=yes

; PRIVILEGES
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; ICONEN (optioneel)
; SetupIconFile=favicon.ico
UninstallDisplayIcon={app}\{#AppExeName}

; WIZARD SETTINGS
WizardStyle=modern

; VERSION INFO
VersionInfoVersion={#AppVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Installatieprogramma
VersionInfoCopyright=Copyright (C) 2025 {#AppPublisher}

[Languages]
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"

[Files]
; MAIN APPLICATION FILES - Relatief pad!
Source: "dist\NVict_Etiketten_Maker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; ICONEN - Optioneel maar aanbevolen voor mooie weergave
; Verwijder de ; als je deze bestanden hebt
 Source: "favicon.ico"; DestDir: "{app}"; Flags: ignoreversion
 Source: "Logo.png"; DestDir: "{app}"; Flags: ignoreversion

; VOORBEELD BESTANDEN naar Documenten\Etiketten Maker (optioneel - uitgecommentarieerd)
; Verwijder de ; om deze bestanden te installeren (als ze bestaan)
 Source: "Voorbeeld_Adressen.xlsx"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist
 Source: "Voorbeeld_Adressen_Uitgebreid.xlsx"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist
 Source: "LEES_MIJ_VOORBEELDEN.txt"; DestDir: "{userdocs}\Etiketten Maker"; Flags: ignoreversion onlyifdoesntexist

[Icons]
; Start Menu snelkoppeling
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Comment: "Excel naar etiketten converteren"

; Desktop snelkoppeling (optioneel)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

; Snelkoppeling voor deinstallatie
Name: "{group}\Verwijder {#AppName}"; Filename: "{uninstallexe}"

; Snelkoppeling naar de voorbeelden map (optioneel - uitgecommentarieerd)
; Name: "{group}\Voorbeeld Bestanden"; Filename: "{userdocs}\Etiketten Maker"; Comment: "Open de map met voorbeeld Excel bestanden"

[Tasks]
Name: "desktopicon"; Description: "Maak een snelkoppeling op het bureaublad"; GroupDescription: "Extra iconen:"; Flags: unchecked

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Start {#AppName}"; Flags: nowait postinstall skipifsilent

[Registry]
; Basisregistratie voor geïnstalleerde applicaties
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#AppVersion}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#AppPublisher}\{#AppName}"; ValueType: string; ValueName: "ExePath"; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey

[Code]
// Uninstall bestaande versie indien aanwezig
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  Result := 0;
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  // Uninstall oude versie voor installatie
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
  
  // Acties na installatie
  if CurStep = ssPostInstall then
  begin
    // Hier kun je extra acties toevoegen na installatie
  end;
end;

// Functie om te controleren of er al een versie geïnstalleerd is
function InitializeSetup(): Boolean;
var
  OldVersion: String;
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

// Functie die wordt uitgevoerd na deinstallatie
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Opruimen van gebruikersdata (optioneel)
    // DelTree(ExpandConstant('{userappdata}\NVict\Etiketten'), True, True, True);
  end;
end;
