# Create_Version_JSON.ps1
# Creates version JSON with proper encoding for release notes
# Voor NVict Etiketten Maker

param(
    [string]$Version,
    [string]$ReleaseNotesFile = "release_notes.txt",
    [string]$OutputFile = "Output\etiketten_version.json"
)

# Lees release notes uit bestand of gebruik fallback
$releaseNotes = "Versie $Version release"

if (Test-Path $ReleaseNotesFile) {
    $releaseNotes = Get-Content $ReleaseNotesFile -Raw -Encoding UTF8
    # Trim whitespace
    $releaseNotes = $releaseNotes.Trim()
}

# Maak JSON object
$json = @{
    version = $Version
    download_url = "https://www.nvict.nl/software/NVict_Etiketten/NVict_Etiketten_Maker_Setup.exe"
    release_notes = $releaseNotes
    release_date = Get-Date -Format "yyyy-MM-dd"
    update_check_url = "https://www.nvict.nl/software/updates/etiketten_version.json"
}

# Converteer naar JSON
$jsonContent = $json | ConvertTo-Json -Depth 10

# Schrijf naar bestand ZONDER BOM (UTF-8 zonder BOM)
# Dit is belangrijk voor JSON parsers
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($OutputFile, $jsonContent, $utf8NoBom)

Write-Host "[OK] JSON created: $OutputFile (UTF-8 without BOM)"
Write-Host ""
Write-Host "Contents:"
Get-Content $OutputFile
