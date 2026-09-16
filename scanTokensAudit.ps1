<#
scanTokensAudit.ps1
But :
- Scanner les JSON d'export du journal d'audit (Org)
- Extraire les événements liés aux tokens (sur base du texte : "token")
- Générer :
  - OrgAudit_TokenEvents_*.csv      (événements filtrés)
  - OrgAudit_ActionsSummary_*.csv   (récap actions)

Entrées :
- Un ou plusieurs fichiers JSON du type OrgAuditEvents_*.json contenant :
  - dataKept : [events]  (si export avec filtrage côté PS)
  OU
  - data     : [events]  (si export brut)

Compatible Windows PowerShell 5.1 / PowerShell ISE.
#>

[CmdletBinding()]
param(
  [string] $BaseDir = $null,
  [string] $JsonPattern = "OrgAuditEvents_*.json"
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
if (-not (Test-Path $exportsDir)) {
  # fallback: JSON à côté du script
  $exportsDir = $BaseDir
}

# ------------------------- Logging simple -------------------------
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $BaseDir ("Scan-AuditJson_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log {
  param([string]$Message)
  $line = "[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Message
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "Scan-AuditJson.ps1 démarré."
Log "Répertoire JSON: $exportsDir"
Log "Motif JSON: $JsonPattern"

# ------------------------- Trouver les JSON -------------------------
$jsonFiles = Get-ChildItem -Path $exportsDir -Filter $JsonPattern -File -ErrorAction SilentlyContinue

if (-not $jsonFiles -or $jsonFiles.Count -eq 0) {
  Log "Aucun fichier JSON trouvé avec le motif '$JsonPattern' dans '$exportsDir'."
  return
}

# FIX: join correct (évite le -join interprété comme paramètre)
$names = $jsonFiles | ForEach-Object { $_.Name }
$namesJoined = [string]::Join(", ", $names)
Log ("JSON trouvés: {0}" -f $namesJoined)

# ------------------------- Structures de sortie -------------------------
$allEvents       = New-Object System.Collections.Generic.List[object]
$tokenEvents     = New-Object System.Collections.Generic.List[object]
$actionsCounters = @{}

function Add-ActionCount([string]$action) {
  if ([string]::IsNullOrWhiteSpace($action)) { $action = "<null>" }
  if (-not $script:actionsCounters.ContainsKey($action)) { $script:actionsCounters[$action] = 0 }
  $script:actionsCounters[$action]++
}

# ------------------------- Lecture & scan -------------------------
foreach ($file in $jsonFiles) {
  Log ("Lecture JSON: {0}" -f $file.FullName)

  $rawContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($rawContent)) {
    Log "  -> Fichier vide, ignoré."
    continue
  }

  try {
    $json = $rawContent | ConvertFrom-Json -ErrorAction Stop
  } catch {
    Log ("  -> Erreur ConvertFrom-Json: {0}" -f $_.Exception.Message)
    continue
  }

  # Selon export:
  $events = @()
  if ($json.dataKept)      { $events = $json.dataKept }
  elseif ($json.dataAll)   { $events = $json.dataAll }   # au cas où ton export contient dataAll/dataKept
  elseif ($json.data)      { $events = $json.data }
  else {
    Log "  -> Ni dataKept, ni dataAll, ni data : rien à scanner."
    continue
  }

  $count = ($events | Measure-Object).Count
  Log ("  -> {0} events trouvés dans ce JSON." -f $count)

  foreach ($e in $events) {
    $allEvents.Add($e) | Out-Null

    $attr   = $e.attributes
    $time   = [string]$attr.time
    $action = [string]$attr.action
    Add-ActionCount -action $action

    $actorEmail = $null
    $actorName  = $null
    if ($attr.actor) {
      $actorEmail = [string]$attr.actor.email
      $actorName  = [string]$attr.actor.name
    }

    $message = $null
    if ($attr -and ($attr.PSObject.Properties.Name -contains "message")) {
      $message = [string]$attr.message
    }

    # Filtre token (simple, modifiable)
    $isTokenRelated = $false
    if ($action  -match "(?i)\btoken\b")   { $isTokenRelated = $true }
    elseif ($message -match "(?i)\btoken\b") { $isTokenRelated = $true }

    if ($isTokenRelated) {
      $tokenEvents.Add([pscustomobject]@{
        Time       = $time
        Action     = $action
        Message    = $message
        ActorEmail = $actorEmail
        ActorName  = $actorName
        SourceFile = $file.Name
      }) | Out-Null
    }
  }
}

Log ("Events totaux (tous JSON confondus)   : {0}" -f $allEvents.Count)
Log ("Events liés à 'token' (action/message): {0}" -f $tokenEvents.Count)

# ------------------------- Export CSV : TokenEvents -------------------------
$tokenCsv = Join-Path $exportsDir ("OrgAudit_TokenEvents_{0}.csv" -f $ts)
$tokenEvents |
  Sort-Object Time |
  Export-Csv -Path $tokenCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Token events exportés dans : {0}" -f $tokenCsv)

# ------------------------- Export CSV : ActionsSummary -------------------------
$summaryCsv = Join-Path $exportsDir ("OrgAudit_ActionsSummary_{0}.csv" -f $ts)

$summary = $actionsCounters.GetEnumerator() |
  Sort-Object Key |
  ForEach-Object {
    [pscustomobject]@{
      Action = $_.Key
      Count  = $_.Value
    }
  }

$summary | Export-Csv -Path $summaryCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log ("Résumé des actions exporté dans : {0}" -f $summaryCsv)

Log "Scan terminé."