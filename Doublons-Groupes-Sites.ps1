<#
.SYNOPSIS
  Doublons-Groupes-Sites.ps1
  Liste et supprime les groupes presents a la fois dans mutexfr et Jiradot.

.DESCRIPTION
  Ce script :
  1. Liste tous les groupes de mutexfr.atlassian.net
  2. Liste tous les groupes de jiradot.atlassian.net
  3. Identifie les groupes en commun (doublons inter-sites)
  4. Mode Dry   : affiche la liste des doublons (aucune action)
  5. Mode Execute : supprime les doublons de Jiradot apres confirmation

  CREDENTIALS :
    secrets\site-admin.xml     => Jiradot (existant)
    secrets\site-mutexfr.xml   => mutexfr (cree au 1er lancement si absent)

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Doublons-Groupes-Sites.ps1                    # Mode Dry (defaut)
  .\Doublons-Groupes-Sites.ps1 -Mode Execute      # Suppression effective
#>

[CmdletBinding()]
param(
  [ValidateSet("Dry","Execute")]
  [string] $Mode = "Dry",

  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("DoublonsGroupes_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("DoublonsGroupes_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# -------------------- Network --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# -------------------- HTTP helper --------------------
function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{ Method=$Method; Uri=$Url; Headers=$Headers; UseBasicParsing=$true; ErrorAction="Stop" }
      if ($Body) { $params["ContentType"]="application/json; charset=utf-8"; $params["Body"]=[System.Text.Encoding]::UTF8.GetBytes($Body) }
      $resp = Invoke-WebRequest @params
      $contentUtf8 = $resp.Content
      try {
        $stream = $resp.RawContentStream
        if ($stream -and $stream.CanSeek) {
          $stream.Position = 0
          $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
          $contentUtf8 = $reader.ReadToEnd(); $reader.Close()
        }
      } catch {
        try {
          $isoBytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($resp.Content)
          $contentUtf8 = [System.Text.Encoding]::UTF8.GetString($isoBytes)
        } catch {}
      }
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$contentUtf8 }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $rd = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $rd.ReadToEnd(); $rd.Close()
      } catch {}
      if ($attempt -gt $MaxRetries) { return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody } }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(5, $attempt)))
        Log ("Retry {0} status={1} in {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
        Start-Sleep -Seconds $sleepSec; continue
      }
      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# -------------------- Load credentials --------------------
function Load-SiteCredentials([string]$CredFile, [string]$SiteName) {
  if (-not (Test-Path $CredFile)) {
    Log "Fichier $CredFile introuvable, creation interactive..." "WARN"
    [System.Windows.Forms.MessageBox]::Show(
      ("Credentials pour {0} manquants.`n`nVous allez fournir :`n  1. URL du site`n  2. Email admin`n  3. API Token" -f $SiteName),
      "Credentials $SiteName", [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl = Read-Host "URL du site (ex: $SiteName.atlassian.net)"
    $inputUrl = $inputUrl -replace "^https?://", "" -replace "/.*$", "" -replace "/$", ""
    $adminEmail = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $CredFile
    Log "Credentials sauvegardes dans $CredFile"
  }

  $data = Import-Clixml -Path $CredFile
  $url = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
    Name    = $SiteName
  }
}

# -------------------- List all groups from a site --------------------
function Get-AllGroups([hashtable]$Site) {
  $groups = New-Object System.Collections.Generic.List[object]
  $startAt = 0; $pageSize = 100

  while ($true) {
    $url = "{0}/rest/api/3/group/bulk?maxResults={1}&startAt={2}" -f $Site.BaseUrl, $pageSize, $startAt
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers

    if (-not $resp.ok) {
      Log ("  Erreur listing groupes {0} : status={1}" -f $Site.Name, $resp.status) "ERROR"
      break
    }

    $json = $resp.content | ConvertFrom-Json

    foreach ($g in $json.values) {
      $groups.Add(@{
        name    = [string]$g.name
        groupId = [string]$g.groupId
      }) | Out-Null
    }

    if ($json.isLast -eq $true -or $groups.Count -ge $json.total) { break }
    $startAt += $pageSize
    Start-Sleep -Milliseconds 200
  }

  Log ("  {0} : {1} groupes trouves" -f $Site.Name, $groups.Count)
  return $groups
}

# ==================== MAIN ====================

Log "================================================================"
Log "  DOUBLONS GROUPES INTER-SITES"
Log ("  Mode: {0}" -f $Mode)
Log "================================================================"

# Charger les credentials
$siteJiradot = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
$siteMutexfr = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-mutexfr.xml") -SiteName "mutexfr"

Log ("Jiradot : {0}" -f $siteJiradot.BaseUrl)
Log ("mutexfr : {0}" -f $siteMutexfr.BaseUrl)

# ==================== ETAPE 1 : Lister les groupes ====================
Log "=== ETAPE 1 : Listing des groupes ==="
$groupsMutexfr = Get-AllGroups -Site $siteMutexfr
$groupsJiradot = Get-AllGroups -Site $siteJiradot

# ==================== ETAPE 2 : Trouver les doublons ====================
Log "=== ETAPE 2 : Identification des doublons ==="

# Creer un HashSet des noms mutexfr pour lookup rapide
$mutexfrNames = New-Object System.Collections.Generic.HashSet[string]
foreach ($g in $groupsMutexfr) { [void]$mutexfrNames.Add($g.name) }

# Trouver les groupes Jiradot qui existent aussi dans mutexfr
$doublons = New-Object System.Collections.Generic.List[object]
foreach ($g in $groupsJiradot) {
  if ($mutexfrNames.Contains($g.name)) {
    $doublons.Add($g) | Out-Null
  }
}

Log ("  Groupes mutexfr : {0}" -f $groupsMutexfr.Count)
Log ("  Groupes Jiradot : {0}" -f $groupsJiradot.Count)
Log ("  Doublons (presents dans les 2 sites) : {0}" -f $doublons.Count)

# Groupes uniquement dans Jiradot (pour info)
$uniqueJiradot = New-Object System.Collections.Generic.List[object]
foreach ($g in $groupsJiradot) {
  if (-not $mutexfrNames.Contains($g.name)) {
    $uniqueJiradot.Add($g) | Out-Null
  }
}
Log ("  Groupes uniquement dans Jiradot : {0}" -f $uniqueJiradot.Count)

if ($doublons.Count -eq 0) {
  Log "Aucun doublon trouve. Rien a faire."
  return
}

# ==================== ETAPE 3 : Export CSV ====================
Log "=== ETAPE 3 : Export des doublons ==="

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvHeader = '"GroupName";"GroupId";"PresentDans";"Action"'
[System.IO.File]::WriteAllText($csvFile, "$csvHeader`r`n", $utf8Bom)

foreach ($g in ($doublons | Sort-Object { $_.name })) {
  $action = if ($Mode -eq "Execute") { "A supprimer de Jiradot" } else { "Doublon (dry)" }
  $line = '"{0}";"{1}";"mutexfr + Jiradot";"{2}"' -f ($g.name -replace '"','""'), $g.groupId, $action
  [System.IO.File]::AppendAllText($csvFile, "$line`r`n", $utf8Bom)
}

Log "  CSV -> $csvFile"

# ==================== ETAPE 4 : Affichage / Suppression ====================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  {0} GROUPES EN DOUBLON" -f $doublons.Count) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

foreach ($g in ($doublons | Sort-Object { $_.name })) {
  Write-Host ("  {0}" -f $g.name) -ForegroundColor Yellow
}

Write-Host ""

if ($Mode -eq "Dry") {
  Write-Host "========================================" -ForegroundColor Green
  Write-Host "  MODE DRY : aucune suppression" -ForegroundColor Green
  Write-Host "  Relancez avec -Mode Execute pour supprimer" -ForegroundColor Green
  Write-Host "========================================" -ForegroundColor Green
  Log "Mode Dry : aucune action effectuee."
}
else {
  # Mode Execute : confirmation avant suppression
  Write-Host "========================================" -ForegroundColor Red
  Write-Host ("  MODE EXECUTE : {0} groupes seront supprimes de Jiradot" -f $doublons.Count) -ForegroundColor Red
  Write-Host "========================================" -ForegroundColor Red
  Write-Host ""

  $confirm = Read-Host "Confirmez la suppression (tapez OUI en majuscules)"

  if ($confirm -ne "OUI") {
    Log "Suppression annulee par l'utilisateur."
    Write-Host "Annule." -ForegroundColor Yellow
  }
  else {
    Log "=== ETAPE 5 : Suppression des doublons de Jiradot ==="
    $cOk = 0; $cErr = 0

    foreach ($g in ($doublons | Sort-Object { $_.name })) {
      $gName = $g.name
      $gId = $g.groupId

      # Suppression par groupId (plus fiable que par nom)
      $url = "{0}/rest/api/3/group?groupId={1}" -f $siteJiradot.BaseUrl, [System.Uri]::EscapeDataString($gId)
      $resp = Invoke-ApiCall -Method "DELETE" -Url $url -Headers $siteJiradot.Headers

      if ($resp.ok -or $resp.status -eq 200 -or $resp.status -eq 204) {
        Log ("  OK : {0} (id={1})" -f $gName, $gId)
        Write-Host ("  OK : {0}" -f $gName) -ForegroundColor Green
        $cOk++
      }
      else {
        Log ("  ERREUR : {0} status={1} {2}" -f $gName, $resp.status, $resp.error) "ERROR"
        Write-Host ("  ERREUR : {0} (status={1})" -f $gName, $resp.status) -ForegroundColor Red
        $cErr++
      }

      Start-Sleep -Milliseconds 300
    }

    Log ("Suppression terminee : {0} OK, {1} erreurs" -f $cOk, $cErr)
  }
}

# ==================== RESUME ====================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  mutexfr         : {0} groupes" -f $groupsMutexfr.Count)
Write-Host ("  Jiradot         : {0} groupes" -f $groupsJiradot.Count)
Write-Host ("  Doublons        : {0}" -f $doublons.Count)
Write-Host ("  Uniques Jiradot : {0}" -f $uniqueJiradot.Count)
Write-Host ("  Mode            : {0}" -f $Mode)
Write-Host ("  CSV             : {0}" -f $csvFile)
Write-Host ("  LOG             : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."