[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Globalization | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

function Pick-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $null = $dlg.ShowDialog()
  $dlg.FileName
}

function Parse-DateUtc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  try { [DateTimeOffset]::Parse($s, $culture).UtcDateTime } catch { $null }
}

$usersCsv = Pick-File "Sélectionner UsersSansApps.csv"
$auditCsv = Pick-File "Sélectionner l'export Audit log (CSV)"
if (-not $usersCsv -or -not $auditCsv) { throw "Fichiers non sélectionnés." }

# Users (delimiter auto ; ou ,)
$h = Get-Content -Path $usersCsv -TotalCount 1 -Encoding UTF8
$delimUsers = ','
if ($h -match ';') { $delimUsers = ';' }
$users = Import-Csv -Path $usersCsv -Delimiter $delimUsers

# Audit (tu as confirmé que c'est ",")
$audit = Import-Csv -Path $auditCsv -Delimiter ','

$target = New-Object System.Collections.Generic.HashSet[string]
foreach ($u in $users) {
  $em = ([string]$u.Email).Trim().ToLowerInvariant()
  if ($em) { [void]$target.Add($em) }
}

# Filtrer les lignes audit qui pointent sur TES users (User Email rempli)
$rows = foreach ($r in $audit) {
  $um = ([string]$r.'User Email').Trim().ToLowerInvariant()
  if (-not $um) { continue }
  if (-not $target.Contains($um)) { continue }

  $dt = Parse-DateUtc ([string]$r.Date)
  [pscustomobject]@{
    DateUtc   = $dt
    UserEmail = $um
    Action    = [string]$r.Action
    Activity  = [string]$r.Activity
    Data      = [string]$r.Data
  }
}

Write-Host ("Audit rows for target users: {0}" -f @($rows).Count)

# 1) Stat par Action/Activity
$byAction = $rows |
  Group-Object Action, Activity |
  Sort-Object Count -Descending |
  Select-Object Count, Name

# 2) Première date vue par user (proxy "arrivée")
$firstSeen = $rows |
  Where-Object { $_.DateUtc } |
  Group-Object UserEmail |
  ForEach-Object {
    $min = ($_.Group | Sort-Object DateUtc | Select-Object -First 1)
    [pscustomobject]@{ UserEmail = $_.Name; FirstSeenUtc = $min.DateUtc.ToString("o"); Action = $min.Action; Activity = $min.Activity }
  } |
  Sort-Object FirstSeenUtc

$ScriptDir  = Split-Path -Parent $PSCommandPath
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"

$byAction | Export-Csv -Path (Join-Path $ExportsDir "Audit_TargetUsers_ByAction_$ts.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
$firstSeen | Export-Csv -Path (Join-Path $ExportsDir "Audit_TargetUsers_FirstSeen_$ts.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Host "Exports -> $ExportsDir"