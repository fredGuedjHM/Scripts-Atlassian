<#
.SYNOPSIS
  Install-ImportExcel.ps1 — Télécharge et installe ImportExcel sans droits admin
  Utilise le proxy système (comme mvp-reporting_v2.ps1)
#>

# Proxy
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
    $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
    $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

$moduleVersion = "7.8.10"
$downloadUrl   = "https://psg-prod-eastus.azureedge.net/packages/importexcel.$moduleVersion.nupkg"
$downloadPath  = "$env:TEMP\ImportExcel.$moduleVersion.nupkg"
$moduleDest    = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\ImportExcel\$moduleVersion"

Write-Host "=== Installation ImportExcel $moduleVersion ===" -ForegroundColor Cyan

# 1. Télécharger
Write-Host "  Téléchargement depuis PSGallery CDN..." -ForegroundColor DarkGray
try {
    $wc = New-Object System.Net.WebClient
    $wc.Proxy = [System.Net.WebRequest]::DefaultWebProxy
    $wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    $wc.DownloadFile($downloadUrl, $downloadPath)
    Write-Host "  ✅ Téléchargé : $downloadPath" -ForegroundColor Green
} catch {
    Write-Host "  ❌ Échec téléchargement : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  → Téléchargez manuellement depuis un navigateur :" -ForegroundColor Yellow
    Write-Host "    https://www.powershellgallery.com/packages/ImportExcel/$moduleVersion" -ForegroundColor Yellow
    Write-Host "    Cliquez 'Manual Download', sauvegardez dans $env:TEMP" -ForegroundColor Yellow
    exit
}

# 2. Extraire (un .nupkg est un .zip)
Write-Host "  Extraction..." -ForegroundColor DarkGray
if (Test-Path $moduleDest) { Remove-Item $moduleDest -Recurse -Force }
New-Item -ItemType Directory -Path $moduleDest -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($downloadPath, $moduleDest)

# Nettoyer les fichiers NuGet inutiles
Remove-Item "$moduleDest\_rels" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$moduleDest\package" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$moduleDest\`[Content_Types`].xml" -Force -ErrorAction SilentlyContinue
Remove-Item "$moduleDest\*.nuspec" -Force -ErrorAction SilentlyContinue

Write-Host "  ✅ Installé dans : $moduleDest" -ForegroundColor Green

# 3. Test
Write-Host "  Test Import-Module..." -ForegroundColor DarkGray
try {
    Import-Module ImportExcel -ErrorAction Stop
    $ver = (Get-Module ImportExcel).Version
    Write-Host "  ✅ ImportExcel $ver chargé avec succès !" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Vous pouvez relancer mvp-reporting_v2.ps1" -ForegroundColor Cyan
} catch {
    Write-Host "  ❌ Échec chargement : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Le script mvp-reporting utilisera le générateur XLSX natif en fallback" -ForegroundColor Yellow
}

# Nettoyage
Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue