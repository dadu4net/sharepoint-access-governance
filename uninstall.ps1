# ==============================================================================
# Noppa SharePoint & Teams Access Governance - De-installatie Script
# ==============================================================================

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host "  Noppa SharePoint & Teams Governance de-installeren      " -ForegroundColor White
Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host ""

$installDir = Join-Path $env:LOCALAPPDATA "Programs\NoppaGovernance"
$startMenuPrograms = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Noppa Solutions"
$desktopFolder = [Environment]::GetFolderPath('Desktop')
$desktopLnk = Join-Path $desktopFolder "Noppa SharePoint & Teams Governance.lnk"

# 1. Sluit actieve instanties
$proc = Get-Process -Name "Noppa-SharePoint-Governance" -ErrorAction SilentlyContinue
if ($proc) {
    Write-Host "[*] Actief proces afsluiten..." -ForegroundColor Gray
    $proc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# 2. Verwijder snelkoppelingen
if (Test-Path $desktopLnk) {
    Remove-Item $desktopLnk -Force
    Write-Host "[-] Bureaublad snelkoppeling verwijderd" -ForegroundColor Gray
}

if (Test-Path $startMenuPrograms) {
    Remove-Item $startMenuPrograms -Recurse -Force
    Write-Host "[-] Startmenu map verwijderd" -ForegroundColor Gray
}

# 3. Verwijder programmabestanden
if (Test-Path $installDir) {
    Remove-Item $installDir -Recurse -Force
    Write-Host "[-] Applicatiemap ($installDir) verwijderd" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Noppa SharePoint & Teams Access Governance is volledig verwijderd." -ForegroundColor Green
Write-Host ""
