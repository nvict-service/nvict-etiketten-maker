# ===================================================
# Upload_FTP_Etiketten.ps1
# FTP Upload voor NVict Etiketten Maker (PowerShell native)
# Geen WinSCP vereist
# ===================================================

param(
    [string]$FtpServer      = "ftp.nvict.nl",
    [string]$FtpUser        = "softwareupload@nvict.nl",
    [string]$FtpPass        = "",
    [string]$FtpPath        = "/NVict_Etiketten",
    [string]$FtpUpdatesPath = "/software/updates",
    [string]$SetupFile      = "Output\NVict_Etiketten_Maker_Setup.exe",
    [string]$SetupName      = "NVict_Etiketten_Maker_Setup.exe"
)

Write-Host ""
Write-Host "========================================"
Write-Host " NVict Etiketten - FTP Upload"
Write-Host "========================================"
Write-Host ""

# Wachtwoord ophalen als niet meegegeven
if ([string]::IsNullOrEmpty($FtpPass)) {
    # Probeer ftp_config.ini te lezen
    if (Test-Path "ftp_config.ini") {
        foreach ($line in Get-Content "ftp_config.ini") {
            if ($line -match "^password=(.+)$") {
                $FtpPass = $Matches[1].Trim()
                Write-Host "[i] Wachtwoord geladen uit ftp_config.ini"
            }
        }
    }
}

if ([string]::IsNullOrEmpty($FtpPass)) {
    $securePass = Read-Host "FTP wachtwoord" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
    $FtpPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

if ([string]::IsNullOrEmpty($FtpPass)) {
    Write-Host "[!] Geen wachtwoord - upload gestopt" -ForegroundColor Red
    exit 1
}

$ftpBase = "ftp://$FtpServer"

function Upload-File {
    param([string]$LocalFile, [string]$RemoteUrl)

    if (-not (Test-Path $LocalFile)) {
        Write-Host "[i] Bestand niet gevonden: $LocalFile - overgeslagen"
        return $true
    }

    $fileName = Split-Path $LocalFile -Leaf
    $fileSize = [math]::Round((Get-Item $LocalFile).Length / 1MB, 2)
    Write-Host "Uploaden: $fileName ($fileSize MB) -> $RemoteUrl"

    try {
        $request = [System.Net.FtpWebRequest]::Create($RemoteUrl)
        $request.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
        $request.Credentials = New-Object System.Net.NetworkCredential($FtpUser, $FtpPass)
        $request.UseBinary = $true
        $request.UsePassive = $true
        $request.KeepAlive = $false

        $fileContent = [System.IO.File]::ReadAllBytes((Resolve-Path $LocalFile))
        $request.ContentLength = $fileContent.Length

        $stream = $request.GetRequestStream()
        $stream.Write($fileContent, 0, $fileContent.Length)
        $stream.Close()

        $response = $request.GetResponse()
        Write-Host "[OK] $fileName geupload" -ForegroundColor Green
        $response.Close()
        return $true
    }
    catch {
        Write-Host "[!] Upload MISLUKT: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

$succes = $true

# Setup installer uploaden
if (-not [string]::IsNullOrEmpty($SetupFile) -and (Test-Path $SetupFile)) {
    $remoteUrl = "$ftpBase$FtpPath/$SetupName"
    $ok = Upload-File -LocalFile $SetupFile -RemoteUrl $remoteUrl
    if (-not $ok) { $succes = $false }
} else {
    Write-Host "[i] Geen setup bestand gevonden - overgeslagen"
}

# Version JSON uploaden
$jsonFile = "Output\etiketten_version.json"
if (Test-Path $jsonFile) {
    $remoteUrl = "$ftpBase$FtpUpdatesPath/etiketten_version.json"
    $ok = Upload-File -LocalFile $jsonFile -RemoteUrl $remoteUrl
    if (-not $ok) { $succes = $false }
} else {
    Write-Host "[i] etiketten_version.json niet gevonden - overgeslagen"
}

# Wachtwoord uit geheugen wissen
if ($BSTR) {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
}

Write-Host ""
if ($succes) {
    Write-Host "========================================"
    Write-Host " UPLOAD VOLTOOID" -ForegroundColor Green
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Beschikbaar op:"
    Write-Host "  Setup:   https://www.nvict.nl/software/NVict_Etiketten/$SetupName"
    Write-Host "  Version: https://www.nvict.nl/software/updates/etiketten_version.json"
    exit 0
} else {
    Write-Host "========================================"
    Write-Host " UPLOAD MISLUKT" -ForegroundColor Red
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Controleer:"
    Write-Host "  - FTP inloggegevens (ftp_config.ini)"
    Write-Host "  - Internetverbinding"
    Write-Host "  - Server bereikbaar"
    exit 1
}
