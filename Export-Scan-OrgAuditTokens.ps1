<#
Export-Scan-OrgAuditTokens.ps1 (ISE-friendly)
Approche:
- NE PAS passer from/to à /events-stream (contourne ADMIN-400-5 Invalid date)
- Télécharger les events les plus récents et filtrer côté PowerShell sur attributes.time (dernier mois)
Docs:
- https://developer.atlassian.com/cloud/admin/organization/rest/intro/#authentication
- https://developer.atlassian.com/cloud/admin/organization/rest/api-group-events/
#>

[CmdletBinding()]
param(
  [string] $OrgId = $null,
  [int]    $DaysBack = 0,      # 0 => prompt (défaut 30)
  [int]    $PageSize = 100,
  [int]    $MaxPages = 50,     # garde-fou (évite de paginer sans fin)
  [string] $BaseDir  = $null
)

try { Set-StrictMode -Off } catch {}

Add-Type -AssemblyName "System.Globalization" | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# ------------------------- Resolve BaseDir -------------------------
$ScriptDir = $BaseDir
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
  if ($PSCommandPath) {
    $ScriptDir = Split-Path -Parent $PSCommandPath
  } elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  }
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = (Get-Location).Path }

# ------------------------- Logging + Export paths -------------------------
$global:LogFile = $null
function Ensure-Dir([string]$path) { if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null } }

function Initialize-Paths {
  param([string]$BaseDir)
  $logsDir = Join-Path $BaseDir "logs"
  $expDir  = Join-Path $BaseDir "exports"
  Ensure-Dir $logsDir
  Ensure-Dir $expDir
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $log    = Join-Path $logsDir ("Export-Scan-OrgAuditTokens_{0}.log" -f $ts)
  $raw    = Join-Path $expDir  ("OrgAuditEvents_{0}.json"        -f $ts)
  $flat   = Join-Path $expDir  ("OrgAuditTokensFlat_{0}.csv"     -f $ts)
  $byUser = Join-Path $expDir  ("OrgAuditTokensByUser_{0}.csv"   -f $ts)
  Set-Content -Path $log -Value "" -Encoding UTF8
  return @{ Log = $log; Raw = $raw; Flat = $flat; ByUser = $byUser }
}

function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$ts][$Level] $Message"
  Write-Host $line
  if ($global:LogFile) { Add-Content -Path $global:LogFile -Value $line -Encoding UTF8 }
}

$paths = Initialize-Paths -BaseDir $ScriptDir
$global:LogFile = $paths.Log

Write-Log "BaseDir:   $ScriptDir" "INFO"
Write-Log "LogFile:   $($paths.Log)" "INFO"
Write-Log "RawJson:   $($paths.Raw)" "INFO"
Write-Log "FlatCsv:   $($paths.Flat)" "INFO"
Write-Log "ByUserCsv: $($paths.ByUser)" "INFO"

# ------------------------- Network prerequisites (TLS + Proxy) -------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
  Write-Log "Proxy système initialisé (DefaultWebProxy)." "INFO"
} catch {
  Write-Log "Impossible d'initialiser le proxy système: $($_.Exception.Message)" "WARN"
}

# ------------------------- Helpers -------------------------
function Get-SafeProp {
  param([Parameter(Mandatory=$true)]$Object, [Parameter(Mandatory=$true)][string]$PropName)
  if ($null -eq $Object) { return $null }
  $p = $Object.PSObject.Properties[$PropName]
  if ($p) { return $p.Value }
  return $null
}

function Get-SafeNested {
  param([Parameter(Mandatory=$true)]$Object, [Parameter(Mandatory=$true)][string[]]$Path)
  $cur = $Object
  foreach ($k in $Path) {
    if ($null -eq $cur) { return $null }
    $cur = Get-SafeProp -Object $cur -PropName $k
  }
  return $cur
}

function Invoke-AdminApiGet {
  param([Parameter(Mandatory=$true)][string]$Url, [Parameter(Mandatory=$true)][hashtable]$Headers)
  try {
    return Invoke-RestMethod -Method GET -Uri $Url -Headers $Headers -ErrorAction Stop
  } catch {
    $details = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = "$details | $($_.ErrorDetails.Message)" }
    throw "Erreur Admin API (GET $Url): $details"
  }
}

function Get-CursorFromLinksNext([string]$NextUrl) {
  if ([string]::IsNullOrWhiteSpace($NextUrl)) { return $null }
  if ($NextUrl -match "cursor=([^&]+)") { return $Matches[1] }
  return $null
}

function Try-GetServerNowUtc {
  param([hashtable]$Headers)

  # Essaie de lire l'heure serveur via header Date (RFC1123) sur /orgs
  $rh = $null
  try {
    $null = Invoke-WebRequest -Method GET `
      -Uri "https://api.atlassian.com/admin/v1/orgs" `
      -Headers $Headers `
      -UseBasicParsing `
      -ResponseHeadersVariable rh `
      -ErrorAction Stop

    $dateHeader = $null
    try { $dateHeader = $rh["Date"] } catch {}
    if ($dateHeader -is [array]) { $dateHeader = $dateHeader[0] }
    if ([string]::IsNullOrWhiteSpace([string]$dateHeader)) { return $null }

    $dto = [DateTimeOffset]::ParseExact([string]$dateHeader, "r", $culture)
    return $dto.UtcDateTime
  } catch {
    return $null
  }
}

function Parse-EventTimeUtc {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  try {
    # attributes.time ressemble à 2025-02-27T18:50:12.281Z
    return [DateTimeOffset]::Parse($s, $culture).UtcDateTime
  } catch {
    return $null
  }
}

# ------------------------- Prompts (ISE-friendly) -------------------------
if ([string]::IsNullOrWhiteSpace($OrgId)) {
  Write-Host ""
  $OrgId = Read-Host -Prompt "OrgId (admin.atlassian.com/o/<OrgId>/...)"
}

if ($DaysBack -le 0) {
  Write-Host ""
  $defaultDays = 30
  $s = Read-Host -Prompt "Nombre de jours à remonter (défaut=$defaultDays)"
  if ([string]::IsNullOrWhiteSpace($s)) { $DaysBack = $defaultDays }
  else {
    $parsed = 0
    if ([int]::TryParse($s, [ref]$parsed) -and $parsed -gt 0) { $DaysBack = $parsed }
    else { Write-Log "Valeur invalide '$s' -> défaut=$defaultDays" "WARN"; $DaysBack = $defaultDays }
  }
}

Write-Host ""
Write-Host "Colle maintenant l'Organization API key (Admin) depuis ton gestionnaire de mots de passe."
$apiKeySecure = Read-Host -Prompt "Organization API key (Bearer)" -AsSecureString
$apiKeyPlain  = [System.Net.NetworkCredential]::new("", $apiKeySecure).Password
if ([string]::IsNullOrWhiteSpace($apiKeyPlain)) { throw "API key vide. Abandon." }

$headers = @{
  Authorization = "Bearer $apiKeyPlain"
  Accept        = "application/json"
}

# ------------------------- Auth sanity check + Server time -------------------------
Write-Log "Test auth: GET https://api.atlassian.com/admin/v1/orgs" "INFO"
$orgs = Invoke-AdminApiGet -Url "https://api.atlassian.com/admin/v1/orgs" -Headers $headers
Write-Log ("Auth OK. Orgs visibles: {0}" -f (($orgs.data | Measure-Object).Count)) "OK"

$serverNowUtc = Try-GetServerNowUtc -Headers $headers
$nowUtc = if ($serverNowUtc) { $serverNowUtc } else { (Get-Date).ToUniversalTime() }

Write-Log ("NowUtc utilisé: {0} (source: {1})" -f $nowUtc.ToString("o"), $(if($serverNowUtc){"server"}else{"local"})) "INFO"

$cutoffUtc = $nowUtc.AddDays(-$DaysBack)
Write-Log ("Cutoff UTC: {0} (dernier {1} jours)" -f $cutoffUtc.ToString("o"), $DaysBack) "INFO"

# ------------------------- Export events-stream (sans from/to) -------------------------
$base = "https://api.atlassian.com/admin/v1/orgs/$OrgId/events-stream"
$urlBase = $base + "?limit=$PageSize"

$allEvents = New-Object System.Collections.Generic.List[object]
$keptEvents = New-Object System.Collections.Generic.List[object]

$page = 0
$cursor = $null
$seenAnyNewerThanCutoff = $false

while ($true) {
  $page++
  if ($MaxPages -gt 0 -and $page -gt $MaxPages) {
    Write-Log "Arrêt: MaxPages atteint ($MaxPages)." "WARN"
    break
  }

  $url = $urlBase
  if (-not [string]::IsNullOrWhiteSpace($cursor)) {
    $url = $url + "&cursor=" + [uri]::EscapeDataString($cursor)
  }

  Write-Log ("GET page {0}: {1}" -f $page, $url) "INFO"
  $resp = Invoke-AdminApiGet -Url $url -Headers $headers

  $data = Get-SafeProp -Object $resp -PropName "data"
  if ($null -eq $data -or $data.Count -eq 0) {
    Write-Log "Aucun event dans cette page." "WARN"
  } else {
    foreach ($e in $data) { $allEvents.Add($e) }

    # Filtrage côté PS sur attributes.time
    $pageTimes = @()
    foreach ($e in $data) {
      $tStr = [string](Get-SafeNested -Object $e -Path @("attributes","time"))
      $tUtc = Parse-EventTimeUtc -s $tStr
      if ($tUtc) { $pageTimes += $tUtc }

      if ($tUtc -and $tUtc -ge $cutoffUtc) {
        $keptEvents.Add($e)
        $seenAnyNewerThanCutoff = $true
      }
    }

    Write-Log ("Page {0}: +{1} events (total={2}) | kept(last {3}d)={4}" -f $page, $data.Count, $allEvents.Count, $DaysBack, $keptEvents.Count) "OK"

    # Stop condition (best-effort):
    # si on a déjà vu des events récents ET que TOUTE la page est plus vieille que cutoff => on peut arrêter.
    if ($seenAnyNewerThanCutoff -and $pageTimes.Count -gt 0) {
      $maxUtc = ($pageTimes | Measure-Object -Maximum).Maximum
      if ($maxUtc -lt $cutoffUtc) {
        Write-Log "Arrêt: la page courante ne contient plus d'events dans la fenêtre demandée." "INFO"
        break
      }
    }
  }

  $nextCursor = Get-SafeNested -Object $resp -Path @("meta","next")
  if ([string]::IsNullOrWhiteSpace([string]$nextCursor)) {
    $nextUrl = Get-SafeNested -Object $resp -Path @("links","next")
    $nextCursor = Get-CursorFromLinksNext ([string]$nextUrl)
  }

  if ([string]::IsNullOrWhiteSpace([string]$nextCursor)) {
    Write-Log "Fin pagination (pas de next cursor)." "INFO"
    break
  }

  $cursor = [string]$nextCursor
}

# ------------------------- Export raw JSON (garde tout + filtré) -------------------------
Write-Log "Export raw JSON..." "INFO"
[pscustomobject]@{
  exportedAt = (Get-Date).ToString("o")
  orgId      = $OrgId
  daysBack   = $DaysBack
  cutoffUtc  = $cutoffUtc.ToString("o")
  countAll   = $allEvents.Count
  countKept  = $keptEvents.Count
  dataAll    = $allEvents
  dataKept   = $keptEvents
} | ConvertTo-Json -Depth 50 | Set-Content -Path $paths.Raw -Encoding UTF8
Write-Log ("Raw JSON exporté -> {0}" -f $paths.Raw) "OK"

# ------------------------- Scan tokens (sur keptEvents uniquement) -------------------------
Write-Log "Scan des keptEvents pour tokenId/tokenLabel..." "INFO"
$flat = New-Object System.Collections.Generic.List[object]

foreach ($e in $keptEvents) {
  $attr = Get-SafeProp -Object $e -PropName "attributes"
  $time   = [string](Get-SafeProp -Object $attr -PropName "time")
  $action = [string](Get-SafeProp -Object $attr -PropName "action")

  $actor = Get-SafeProp -Object $attr -PropName "actor"
  $actorEmail = [string](Get-SafeProp -Object $actor -PropName "email")
  $actorName  = [string](Get-SafeProp -Object $actor -PropName "name")
  $actorId    = [string](Get-SafeProp -Object $actor -PropName "id")

  $auth = Get-SafeProp -Object $actor -PropName "auth"
  $authType   = [string](Get-SafeProp -Object $auth -PropName "authType")
  $tokenId    = [string](Get-SafeProp -Object $auth -PropName "tokenId")
  $tokenLabel = [string](Get-SafeProp -Object $auth -PropName "tokenLabel")

  $obo = Get-SafeProp -Object $actor -PropName "onBehalfOf"
  $oboEmail = [string](Get-SafeProp -Object $obo -PropName "email")
  $oboName  = [string](Get-SafeProp -Object $obo -PropName "name")
  $oboId    = [string](Get-SafeProp -Object $obo -PropName "id")

  if (-not [string]::IsNullOrWhiteSpace($tokenId) -or -not [string]::IsNullOrWhiteSpace($tokenLabel)) {
    $flat.Add([pscustomobject]@{
      EventTime       = $time
      Action          = $action
      ActorEmail      = $actorEmail
      ActorName       = $actorName
      ActorId         = $actorId
      AuthType        = $authType
      TokenId         = $tokenId
      TokenLabel      = $tokenLabel
      OnBehalfOfEmail = $oboEmail
      OnBehalfOfName  = $oboName
      OnBehalfOfId    = $oboId
    })
  }
}

$flat | Sort-Object EventTime | Export-Csv -Path $paths.Flat -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log ("Tokens flat export: {0} lignes -> {1}" -f $flat.Count, $paths.Flat) "OK"

$byUser = $flat |
  Group-Object TokenId, TokenLabel, ActorEmail, OnBehalfOfEmail |
  ForEach-Object {
    $items = $_.Group
    $first = ($items | Sort-Object EventTime | Select-Object -First 1)
    $last  = ($items | Sort-Object EventTime | Select-Object -Last 1)
    [pscustomobject]@{
      TokenId         = $first.TokenId
      TokenLabel      = $first.TokenLabel
      AuthType        = $first.AuthType
      ActorEmail      = $first.ActorEmail
      OnBehalfOfEmail = $first.OnBehalfOfEmail
      FirstSeen       = $first.EventTime
      LastSeen        = $last.EventTime
      EventsCount     = $items.Count
    }
  }

$byUser | Sort-Object ActorEmail, TokenLabel, TokenId | Export-Csv -Path $paths.ByUser -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log ("TokensByUser export: {0} lignes -> {1}" -f ($byUser | Measure-Object).Count, $paths.ByUser) "OK"

Write-Log "Terminé." "OK"