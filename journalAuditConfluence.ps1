# Configuration proxy entreprise
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

$siteData = Import-Clixml "C:\Users\GUEDJ-F\OneDrive - Harmonie Mutuelle\Documents\powershell\secrets\site-admin.xml"
$siteUrl = [string]$siteData.SiteUrl
$siteEmail = [string]$siteData.Email
$siteToken = [System.Net.NetworkCredential]::new("", $siteData.ApiTokenSecureString).Password
$auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${siteEmail}:${siteToken}"))
$h = @{Authorization="Basic $auth";Accept="application/json"}
$proxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy("https://$siteUrl")

# Rechercher toutes les categories d'evenements sur les 30 derniers jours (1000 max)
Write-Host "=== Scan audit log : 1000 evenements ===" -ForegroundColor Cyan

$startMs = [long]([DateTimeOffset]::UtcNow.AddDays(-30).ToUnixTimeMilliseconds())
$url = "https://$siteUrl/wiki/rest/api/audit?limit=1000&startDate=$startMs"

try {
  $r = Invoke-WebRequest -Uri $url -Headers $h -UseBasicParsing -ErrorAction Stop -ProxyUseDefaultCredentials -Proxy $proxyUri
  $isoBytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($r.Content)
  $utf8 = [System.Text.Encoding]::UTF8.GetString($isoBytes)
  $json = $utf8 | ConvertFrom-Json

  Write-Host ("Total evenements: {0}" -f $json.results.Count) -ForegroundColor Green

  # Compter par category + summary
  Write-Host "`n=== Categories ===" -ForegroundColor Cyan
  $json.results | Group-Object -Property category | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,5}x  {1}" -f $_.Count, $_.Name) -ForegroundColor Yellow
  }

  Write-Host "`n=== Types d'evenements ===" -ForegroundColor Cyan
  $json.results | Group-Object -Property summary | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,5}x  {1}" -f $_.Count, $_.Name) -ForegroundColor Yellow
  }

  # Chercher specifiquement les events avec un space associe
  Write-Host "`n=== Evenements avec espace associe ===" -ForegroundColor Cyan
  $withSpace = $json.results | Where-Object {
    $found = $false
    if ($_.associatedObjects) {
      foreach ($ao in $_.associatedObjects) {
        if ([string]$ao.objectType -match "space") { $found = $true; break }
      }
    }
    $found
  }
  Write-Host ("  {0} evenements sur {1} ont un espace associe" -f $withSpace.Count, $json.results.Count) -ForegroundColor Green

  # Evenements avec espace, groupes par type
  if ($withSpace.Count -gt 0) {
    $withSpace | Group-Object -Property summary | Sort-Object Count -Descending | ForEach-Object {
      Write-Host ("  {0,5}x  {1}" -f $_.Count, $_.Name) -ForegroundColor Yellow
    }
  }

} catch {
  $s = 0; try { $s = [int]$_.Exception.Response.StatusCode } catch {}
  Write-Host "ERREUR status=$s : $($_.Exception.Message)" -ForegroundColor Red
}