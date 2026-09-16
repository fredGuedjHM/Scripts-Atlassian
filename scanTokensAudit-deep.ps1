<#
scanTokensAudit-deep.ps1
But:
- Scanner un/des JSON d'export Org audit (OrgAuditEvents_*.json)
- Produire des CSV de diagnostic + détection multi-mots-clés

Compatible Windows PowerShell 5.1 + PowerShell ISE.
#>

[CmdletBinding()]
param(
  [string] $BaseDir = $null,
  [string] $JsonPattern = "OrgAuditEvents_*.json",
  [string[]] $Keywords = @(
    # Ajuste librement après lecture de ActionsSummary
    "token","api token","api_token","pat","personal access","personal_access",
    "api key","api_key","key","secret","credential","oauth","client","auth","authentication"
  ),
  [int] $MaxKeywordHits = 5000  # garde-fou
)

try { Set-StrictMode -Off } catch {}

Add-Type -AssemblyName "System.Globalization" | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# ------------------------- Resolve BaseDir -------------------------
if ([string]::IsNullOrWhiteSpace($BaseDir)) {
  if ($PSCommandPath) {
    $BaseDir = Split-Path -Parent $PSCommandPath
  } elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
    $BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  } else {
    $BaseDir = (Get-Location).Path
  }
}
Write-Host "BaseDir: $BaseDir"

$exportsDir = Join-Path $BaseDir "exports"
if (-not (Test-Path $exportsDir)) { $exportsDir = $BaseDir }

# ------------------------- Logging -------------------------
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $BaseDir ("Scan-AuditJson-Deep_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$Message) {
  $line = "[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Message
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "scanTokensAudit-deep.ps1 démarré."
Log "Répertoire JSON: $exportsDir"
Log "Motif JSON: $JsonPattern"
Log ("Keywords: {0}" -f ([string]::Join(", ", $Keywords)))

# ------------------------- Find JSON files -------------------------
$jsonFiles = Get-ChildItem -Path $exportsDir -Filter $JsonPattern -File -ErrorAction SilentlyContinue
if (-not $jsonFiles -or $jsonFiles.Count -eq 0) {
  Log "Aucun fichier JSON trouvé."
  return
}
Log ("JSON trouvés: {0}" -f ([string]::Join(", ", ($jsonFiles | ForEach-Object { $_.Name }))))

# ------------------------- Counters + lists -------------------------
$allEvents = New-Object System.Collections.Generic.List[object]
$actionsCounters = @{}
$attrKeysCounters = @{}
$authTypeCounters = @{}
$keywordHits = New-Object System.Collections.Generic.List[object]
$tokenFields = New-Object System.Collections.Generic.List[object]

function Inc([hashtable]$ht, [string]$key) {
  if ([string]::IsNullOrWhiteSpace($key)) { $key = "<null>" }
  if (-not $ht.ContainsKey($key)) { $ht[$key] = 0 }
  $ht[$key]++
}

function Get-AttrKeys([object]$attr) {
  if ($null -eq $attr) { return @() }
  return @($attr.PSObject.Properties | ForEach-Object { $_.Name })
}

function Contains-AnyKeyword([string]$text, [string[]]$keywords, [ref]$matched) {
  $matched.Value = $null
  if ([string]::IsNullOrWhiteSpace($text)) { return $false }
  foreach ($k in $keywords) {
    if ([string]::IsNullOrWhiteSpace($k)) { continue }
    if ($text -match [Regex]::Escape($k)) { $matched.Value = $k; return $true }
    # fallback case-insensitive contains (more permissive)
    if ($text.ToLowerInvariant().Contains($k.ToLowerInvariant())) { $matched.Value = $k; return $true }
  }
  return $false
}

# ------------------------- Read & scan -------------------------
foreach ($file in $jsonFiles) {
  Log ("Lecture JSON: {0}" -f $file.FullName)
  $rawContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($rawContent)) { Log "  -> vide"; continue }

  try { $json = $rawContent | ConvertFrom-Json -ErrorAction Stop }
  catch { Log ("  -> ConvertFrom-Json KO: {0}" -f $_.Exception.Message); continue }

  $events = @()
  if ($json.dataKept)      { $events = $json.dataKept }
  elseif ($json.dataAll)   { $events = $json.dataAll }
  elseif ($json.data)      { $events = $json.data }
  else { Log "  -> Ni dataKept ni dataAll ni data."; continue }

  Log ("  -> {0} events trouvés dans ce JSON." -f (($events | Measure-Object).Count))

  foreach ($e in $events) {
    $allEvents.Add($e) | Out-Null

    $attr = $e.attributes
    $action = $null
    $time   = $null
    $message = $null

    if ($attr) {
      $action = [string]$attr.action
      $time   = [string]$attr.time
      if ($attr.PSObject.Properties.Name -contains "message") { $message = [string]$attr.message }
    }

    Inc $actionsCounters $action

    # keys inventory
    foreach ($k in (Get-AttrKeys $attr)) { Inc $attrKeysCounters $k }

    # authType inventory (si présent)
    $authType = $null
    try { $authType = [string]$attr.actor.auth.authType } catch {}
    if (-not [string]::IsNullOrWhiteSpace($authType)) { Inc $authTypeCounters $authType }

    # token fields (si présents)
    $tokenId = $null; $tokenLabel = $null
    try { $tokenId = [string]$attr.actor.auth.tokenId } catch {}
    try { $tokenLabel = [string]$attr.actor.auth.tokenLabel } catch {}

    if (-not [string]::IsNullOrWhiteSpace($tokenId) -or -not [string]::IsNullOrWhiteSpace($tokenLabel)) {
      $tokenFields.Add([pscustomobject]@{
        Time       = $time
        Action     = $action
        ActorEmail = [string]($attr.actor.email)
        ActorName  = [string]($attr.actor.name)
        AuthType   = $authType
        TokenId    = $tokenId
        TokenLabel = $tokenLabel
        SourceFile = $file.Name
      }) | Out-Null
    }

    # keyword scan (action/message + fallback JSON compressé de attributes)
    if ($keywordHits.Count -lt $MaxKeywordHits) {
      $m = $null
      $hit = $false

      $actionMsg = (($action + "`n" + $message) -as [string])
      if (Contains-AnyKeyword -text $actionMsg -keywords $Keywords -matched ([ref]$m)) {
        $hit = $true
      } else {
        # fallback: scan texte JSON des attributes (peut être plus coûteux)
        $attrJson = $null
        try { $attrJson = ($attr | ConvertTo-Json -Depth 20 -Compress) } catch {}
        if (Contains-AnyKeyword -text $attrJson -keywords $Keywords -matched ([ref]$m)) { $hit = $true }
      }

      if ($hit) {
        $keywordHits.Add([pscustomobject]@{
          Time        = $time
          Action      = $action
          Matched     = $m
          ActorEmail  = [string]($attr.actor.email)
          ActorName   = [string]($attr.actor.name)
          AuthType    = $authType
          Message     = $message
          SourceFile  = $file.Name
        }) | Out-Null
      }
    }
  }
}

Log ("Total events: {0}" -f $allEvents.Count)
Log ("Keyword hits: {0}" -f $keywordHits.Count)
Log ("Token fields hits (tokenId/tokenLabel): {0}" -f $tokenFields.Count)

# ------------------------- Export CSVs -------------------------
$actionsCsv = Join-Path $exportsDir ("OrgAudit_ActionsSummary_{0}.csv" -f $ts)
$actionsCounters.GetEnumerator() |
  Sort-Object Value -Descending |
  ForEach-Object { [pscustomobject]@{ Action = $_.Key; Count = $_.Value } } |
  Export-Csv -Path $actionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Export: {0}" -f $actionsCsv)

$keysCsv = Join-Path $exportsDir ("OrgAudit_AttributeKeysSummary_{0}.csv" -f $ts)
$attrKeysCounters.GetEnumerator() |
  Sort-Object Value -Descending |
  ForEach-Object { [pscustomobject]@{ AttributeKey = $_.Key; Count = $_.Value } } |
  Export-Csv -Path $keysCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Export: {0}" -f $keysCsv)

$authCsv = Join-Path $exportsDir ("OrgAudit_AuthTypesSummary_{0}.csv" -f $ts)
$authTypeCounters.GetEnumerator() |
  Sort-Object Value -Descending |
  ForEach-Object { [pscustomobject]@{ AuthType = $_.Key; Count = $_.Value } } |
  Export-Csv -Path $authCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Export: {0}" -f $authCsv)

$hitsCsv = Join-Path $exportsDir ("OrgAudit_KeywordHits_{0}.csv" -f $ts)
$keywordHits |
  Sort-Object Time |
  Export-Csv -Path $hitsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Export: {0}" -f $hitsCsv)

$tokenFieldsCsv = Join-Path $exportsDir ("OrgAudit_TokenFields_{0}.csv" -f $ts)
$tokenFields |
  Sort-Object Time |
  Export-Csv -Path $tokenFieldsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Export: {0}" -f $tokenFieldsCsv)

Log "Scan terminé."