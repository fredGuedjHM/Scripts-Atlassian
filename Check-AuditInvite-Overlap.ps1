[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Pick-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $null = $dlg.ShowDialog()
  return $dlg.FileName
}

function Import-CsvUsers([string]$path) {
  $h = Get-Content -Path $path -TotalCount 1 -Encoding UTF8
  $delim = ','
  if ($h -match ';') { $delim = ';' }
  return Import-Csv -Path $path -Delimiter $delim
}

$usersPath = Pick-File "Sélectionner UsersSansApps.csv"
$auditPath = Pick-File "Sélectionner l'export Audit log (CSV)"
if (-not $usersPath -or -not $auditPath) { throw "Fichiers non sélectionnés." }

$users = Import-CsvUsers $usersPath
$audit = Import-Csv -Path $auditPath -Delimiter ','

# emails cibles
$target = New-Object System.Collections.Generic.HashSet[string]
foreach ($u in $users) {
  $em = ([string]$u.Email).Trim().ToLowerInvariant()
  if ($em) { [void]$target.Add($em) }
}

# invite-like
$inviteRegex = '(?i)(invite|invited|invitation|invit[eé]|invité|réinvité|re-invited)'
$inviteRows = @()
foreach ($r in $audit) {
  $um = ([string]$r.'User Email').Trim().ToLowerInvariant()
  $txt = ("{0} | {1}" -f $r.Action, $r.Activity)
  if ($txt -match $inviteRegex) {
    $inviteRows += $r
  }
}

# stats
$inviteDistinctUsers = New-Object System.Collections.Generic.HashSet[string]
$overlapDistinctUsers = New-Object System.Collections.Generic.HashSet[string]

foreach ($r in $inviteRows) {
  $um = ([string]$r.'User Email').Trim().ToLowerInvariant()
  if ($um) {
    [void]$inviteDistinctUsers.Add($um)
    if ($target.Contains($um)) { [void]$overlapDistinctUsers.Add($um) }
  }
}

Write-Host ("UsersSansApps distinct emails: {0}" -f $target.Count)
Write-Host ("Audit invite-like rows: {0}" -f $inviteRows.Count)
Write-Host ("Audit invite-like distinct User Email: {0}" -f $inviteDistinctUsers.Count)
Write-Host ("OVERLAP distinct User Email (invite-like ∩ UsersSansApps): {0}" -f $overlapDistinctUsers.Count)

# exports pour inspection
$ScriptDir = Split-Path -Parent $PSCommandPath
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"

$byAction = $inviteRows | Group-Object Action, Activity | Sort-Object Count -Descending |
  Select-Object Count, Name
$byAction | Export-Csv -Path (Join-Path $ExportsDir "Audit_InviteLike_ByAction_$ts.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8

$overlapRows = $inviteRows | Where-Object {
  $target.Contains((([string]$_.'User Email').Trim().ToLowerInvariant()))
}
$overlapRows | Select-Object Date,'User Email',Action,Activity,Data |
  Export-Csv -Path (Join-Path $ExportsDir "Audit_InviteLike_OverlapRows_$ts.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Host "Exports -> $ExportsDir"