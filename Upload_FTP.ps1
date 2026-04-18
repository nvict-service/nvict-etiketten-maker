# ===================================================
# Upload_FTP.ps1 - NVict Etiketten Maker
# Uploadt setup + version JSON naar nvict.nl
# Kan zelfstandig worden uitgevoerd
# ===================================================

Write-Host ""
Write-Host "========================================"
Write-Host " NVict Etiketten Maker - FTP Upload"
Write-Host "========================================"
Write-Host ""

# Controleer setup bestand
$setupFile = $null
if (Test-Path "Output\NVict_Etiketten_Maker_Setup.exe") {
    $setupFile = "Output\NVict_Etiketten_Maker_Setup.exe"
    $setupName = "NVict_Etiketten_Maker_Setup.exe"
    Write-Host "[OK] Setup gevonden: NVict_Etiketten_Maker_Setup.exe"
} else {
    Write-Host "[ ] Geen setup bestand gevonden in Output\" -ForegroundColor Yellow
    Write-Host "    Alleen version JSON wordt geupload (als die bestaat)."
}

if (-not (Test-Path "Output\etiketten_version.json") -and -not $setupFile) {
    Write-Host ""
    Write-Host "[!] Niets om te uploaden. Voer eerst het release script uit." -ForegroundColor Red
    Read-Host "Druk op Enter om af te sluiten"
    exit 1
}

# Standaard FTP instellingen
$ftpServer  = "ftp.nvict.nl"
$ftpUser    = "softwareupload@nvict.nl"
$ftpPath    = "/NVict_Etiketten"
$ftpUpdPath = "/software/updates"
$ftpPass    = $null

# Lees ftp_config.ini indien aanwezig
if (Test-Path "ftp_config.ini") {
    Write-Host "[i] Config laden uit ftp_config.ini..."
    Get-Content "ftp_config.ini" | ForEach-Object {
        if ($_ -match "^server\s*=\s*(.+)$")        { $ftpServer  = $matches[1].Trim() }
        if ($_ -match "^username\s*=\s*(.+)$")       { $ftpUser    = $matches[1].Trim() }
        if ($_ -match "^remote_path\s*=\s*(.+)$")    { $ftpPath    = $matches[1].Trim() }
        if ($_ -match "^updates_path\s*=\s*(.+)$")   { $ftpUpdPath = $matches[1].Trim() }
        if ($_ -match "^password\s*=\s*(.+)$")       { $ftpPass    = $matches[1].Trim() }
    }
}

Write-Host ""
Write-Host "FTP Instellingen:"
Write-Host "  Server:       $ftpServer"
Write-Host "  Gebruiker:    $ftpUser"
Write-Host "  Setup pad:    $ftpPath"
Write-Host "  Updates pad:  $ftpUpdPath"
Write-Host ""

# Wachtwoord ophalen indien niet in config
if (-not $ftpPass) {
    $securePass = Read-Host "Voer FTP wachtwoord in" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
    $ftpPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

# Helper functie voor FTP upload
function Upload-FTPFile {
    param(
        [string]$LocalFile,
        [string]$RemoteUri,
        [string]$User,
        [string]$Pass
    )

    $fullPath = Resolve-Path $LocalFile
    $sizeMB   = [math]::Round((Get-Item $fullPath).Length / 1MB, 2)
    Write-Host "Uploading: $(Split-Path $LocalFile -Leaf) ($sizeMB MB) ..." -NoNewline

    $req = [System.Net.FtpWebRequest]::Create($RemoteUri)
    $req.Method      = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $req.Credentials = New-Object System.Net.NetworkCredential($User, $Pass)
    $req.UseBinary   = $true
    $req.UsePassive  = $true

    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $req.ContentLength = $bytes.Length

    $stream = $req.GetRequestStream()
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Close()

    $resp = $req.GetResponse()
    Write-Host " VOLTOOID" -ForegroundColor Green
    $resp.Close()
}

# Upload uitvoeren
Write-Host "========================================"
Write-Host "Uploaden naar FTP"
Write-Host "========================================"
Write-Host ""

$uploadOK = $true

try {
    if ($setupFile) {
        Upload-FTPFile `
            -LocalFile  $setupFile `
            -RemoteUri  "ftp://$ftpServer$ftpPath/$setupName" `
            -User       $ftpUser `
            -Pass       $ftpPass
    }

    if (Test-Path "Output\etiketten_version.json") {
        Upload-FTPFile `
            -LocalFile  "Output\etiketten_version.json" `
            -RemoteUri  "ftp://$ftpServer$ftpUpdPath/etiketten_version.json" `
            -User       $ftpUser `
            -Pass       $ftpPass
    }
} catch {
    $uploadOK = $false
    Write-Host " MISLUKT" -ForegroundColor Red
    Write-Host ""
    Write-Host "[!] Foutmelding: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Mogelijke oorzaken:"
    Write-Host "  - Verkeerd wachtwoord"
    Write-Host "  - Geen internetverbinding"
    Write-Host "  - Server niet bereikbaar"
    Write-Host "  - Geen schrijfrechten op server"
}

# Geheugen opschonen
if ($BSTR) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR) }

Write-Host ""
if ($uploadOK) {
    Write-Host "========================================"
    Write-Host " UPLOAD VOLTOOID!" -ForegroundColor Green
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Bestanden beschikbaar op:"
    if ($setupFile) {
        Write-Host "  Setup:   https://www.nvict.nl/software/NVict_Etiketten/$setupName"
    }
    if (Test-Path "Output\etiketten_version.json") {
        Write-Host "  Version: https://www.nvict.nl/software/updates/etiketten_version.json"
    }
} else {
    Write-Host "[!] Upload niet volledig voltooid. Controleer de foutmelding hierboven." -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Druk op Enter om af te sluiten"
