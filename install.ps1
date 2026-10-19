# ==============================================================================
# Noppa SharePoint & Teams Access Governance - Installatie Script
# ==============================================================================
# Installeert de standalone Windows Desktop Executable (.exe) naar %LOCALAPPDATA%
# en maakt snelkoppelingen aan in het Startmenu en op het Bureaublad.
# Geen administratorrechten vereist!
# ==============================================================================

[CmdletBinding()]
param (
    [switch]$Standalone, # Gebruik de 100% self-contained variant uit dist\standalone
    [switch]$LaunchAfterInstall
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Noppa SharePoint & Teams Access Governance Installer    " -ForegroundColor White
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Get-Location }

# Bepaal bronbestand
$sourceExe = ""
if ($Standalone) {
    $sourceExe = Join-Path $scriptRoot "dist\standalone\Noppa-SharePoint-Governance.exe"
} else {
    $sourceExe = Join-Path $scriptRoot "dist\Noppa-SharePoint-Governance.exe"
    if (-not (Test-Path $sourceExe)) {
        # Fallback naar standalone als reguliere dist niet aanwezig is
        $sourceExe = Join-Path $scriptRoot "dist\standalone\Noppa-SharePoint-Governance.exe"
    }
}

if (-not (Test-Path $sourceExe)) {
    Write-Host "[!] Fout: Kan 'Noppa-SharePoint-Governance.exe' niet vinden in dist\" -ForegroundColor Red
    Write-Host "    Voer eerst 'npm run build:exe' uit om de executable te genereren." -ForegroundColor Yellow
    exit 1
}

# Installatielocatie
$installDir = Join-Path $env:LOCALAPPDATA "Programs\NoppaGovernance"
Write-Host "[*] Installatielocatie: $installDir" -ForegroundColor Gray

if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

# Bestanden kopiëren
Write-Host "[*] Bestanden kopiëren..." -ForegroundColor Cyan
$sourceDir = Split-Path $sourceExe -Parent
Get-ChildItem -Path $sourceDir -Recurse | Copy-Item -Destination $installDir -Force -Recurse

# Icon veiligstellen
$sourceIco = Join-Path $scriptRoot "desktop\app.ico"
$targetIco = Join-Path $installDir "app.ico"
if (Test-Path $sourceIco) {
    Copy-Item -Path $sourceIco -Destination $targetIco -Force
}

$targetExe = Join-Path $installDir "Noppa-SharePoint-Governance.exe"

# Snelkoppelingen aanmaken via WScript.Shell
$wsh = New-Object -ComObject WScript.Shell

# 1. Startmenu
$startMenuPrograms = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Noppa Solutions"
if (-not (Test-Path $startMenuPrograms)) {
    New-Item -ItemType Directory -Path $startMenuPrograms -Force | Out-Null
}
$startMenuLnk = Join-Path $startMenuPrograms "Noppa SharePoint & Teams Governance.lnk"
$shortcut = $wsh.CreateShortcut($startMenuLnk)
$shortcut.TargetPath = $targetExe
$shortcut.WorkingDirectory = $installDir
$shortcut.Description = "Noppa SharePoint & Teams Access Governance Portal"
if (Test-Path $targetIco) {
    $shortcut.IconLocation = "$targetIco,0"
}
$shortcut.Save()
Write-Host "[+] Startmenu snelkoppeling aangemaakt" -ForegroundColor Green

# 2. Bureaublad
$desktopFolder = [Environment]::GetFolderPath('Desktop')
$desktopLnk = Join-Path $desktopFolder "Noppa SharePoint & Teams Governance.lnk"
$shortcutDesktop = $wsh.CreateShortcut($desktopLnk)
$shortcutDesktop.TargetPath = $targetExe
$shortcutDesktop.WorkingDirectory = $installDir
$shortcutDesktop.Description = "Noppa SharePoint & Teams Access Governance Portal"
if (Test-Path $targetIco) {
    $shortcutDesktop.IconLocation = "$targetIco,0"
}
$shortcutDesktop.Save()
Write-Host "[+] Bureaublad snelkoppeling aangemaakt" -ForegroundColor Green

# 3. Genereer uninstall script in installatiedirectory
$uninstallScript = @"
Write-Host "Noppa SharePoint & Teams Governance verwijderen..." -ForegroundColor Yellow

`$installDir = "$installDir"
`$startMenu = "$startMenuPrograms"
`$desktopLnk = "$desktopLnk"

if (Test-Path `$desktopLnk) { Remove-Item `$desktopLnk -Force }
if (Test-Path `$startMenu) { Remove-Item `$startMenu -Recurse -Force }

# Wacht tot proces eventueel afgesloten is
Get-Process -Name "Noppa-SharePoint-Governance" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

if (Test-Path `$installDir) {
    Remove-Item `$installDir -Recurse -Force
}

Write-Host "Succesvol verwijderd!" -ForegroundColor Green
"@
Set-Content -Path (Join-Path $installDir "uninstall.ps1") -Value $uninstallScript -Encoding UTF8

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Installatie succesvol voltooid!                         " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Applicatie: $targetExe" -ForegroundColor White
Write-Host "  Snelkoppelingen toegevoegd aan Startmenu en Bureaublad." -ForegroundColor Gray
Write-Host ""

if ($LaunchAfterInstall) {
    Write-Host "[*] Applicatie starten..." -ForegroundColor Cyan
    Start-Process -FilePath $targetExe
}
