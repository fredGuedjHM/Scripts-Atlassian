[CmdletBinding()]
param(
  # Chemin vers UsersSansApps.csv (si vide, on demandera)
  [string] $UsersSansAppsCsv = "",

  # Chemin vers l’export audit log CSV 01/01/2025+ (si vide, on demandera)
  [string] $AuditLogCsv = "",

  # Si aucune invite >= PreciseFromYear trouvée => "Avant <PreciseFromYear>"
  [int] $PreciseFromYear = 2025
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Globalization | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

$ScriptDir  = Split-Path -Parent $PSCommandPath
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"

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

function Parse-DateUtc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  try { return [DateTimeOffset]::Parse($s, $culture).UtcDateTime } catch { return $null }
}

function Update-Earliest([hashtable]$map, [string]$key, [datetime]$dt) {
  if (-not $key -or -not $dt) { return }
  if (-not $map.ContainsKey($key)) { $map[$key] = $dt; return }
  if ($dt -lt $map[$key]) { $map[$key] = $dt }
}

# ---- 1) UsersSansApps.csv
if ([string]::IsNullOrWhiteSpace($UsersSansAppsCsv)) {
  $UsersSansAppsCsv = Pick-File "Sélectionner UsersSansApps.csv (avec Email + AccountId)"
}
if (-not (Test-Path $UsersSansAppsCsv)) {
  throw "UsersSansAppsCsv introuvable: $UsersSansAppsCsv"
}

# délimiteur auto ; ou ,
$rawUsers = Get-Content -Path $UsersSansAppsCsv -TotalCount 1 -Encoding UTF8
$delimUsers = ','
if ($rawUsers -match ';') { $delimUsers = ';' }

$users = Import-Csv -Path $UsersSansAppsCsv -Delimiter $delimUsers
if (-not $users -or $users.Count -eq 0) { throw "UsersSansApps.csv vide." }

$targetEmails = New-Object System.Collections.Generic.HashSet[string]
foreach ($u in $users) {
  $em = ([string]$u.Email).Trim().ToLowerInvariant()
  if ($em) { [void]$targetEmails.Add($em) }
}
if ($targetEmails.Count -eq 0) { throw "Aucune colonne Email exploitable dans UsersSansApps.csv." }

Write-Host ("Users loaded={0} distinct Emails={1}" -f $users.Count, $targetEmails.Count)

# ---- 2) AuditLogCsv
if ([string]::IsNullOrWhiteSpace($AuditLogCsv)) {
  $AuditLogCsv = Pick-File "Sélectionner l'export Audit log (2025+) : 2026-05-07T152041Z_audit_log.csv"
}
if (-not (Test-Path $AuditLogCsv)) {
  throw "AuditLogCsv introuvable: $AuditLogCsv"
}

# d’après Inspect-AuditLogCsv : délimiteur = ','
$audit = Import-Csv -Path $AuditLogCsv -Delimiter ','

if (-not $audit -or $audit.Count -eq 0) { throw "Audit log CSV vide." }

# Vérif colonnes attendues
$cols = $audit[0].PSObject.Properties.Name
foreach ($c in @("Date","User Email","Action","Activity")) {
  if (-not ($cols -contains $c)) {
    throw "La colonne '$c' est absente de l'audit log. Colonnes dispos: $($cols -join ', ')"
  }
}

Write-Host ("Audit rows={0}, columns={1}" -f $audit.Count, ($cols -join ", "))

# regex invite sur Action ou Activity
$inviteRegex = "(?i)(invite|invited|invitation|invit[eé]|invité|réinvité|re-invited|ajouté|added)"

$fromYearUtc = [datetime]::SpecifyKind(
  [datetime]::ParseExact("$PreciseFromYear-01-01","yyyy-MM-dd",$culture),
  [DateTimeKind]::Utc
)

$inviteByEmail = @{}
$matchedRows = 0

foreach ($r in $audit) {
  $action   = [string]$r.'Action'
  $activity = [string]$r.'Activity'
  $userMail = ([string]$r.'User Email').Trim().ToLowerInvariant()
  $dateStr  = [string]$r.'Date'

  if (-not $userMail -or -not $targetEmails.Contains($userMail)) { continue }

  $text = "$action | $activity"
  if ($text -notmatch $inviteRegex) { continue }

  $dt = Parse-DateUtc $dateStr
  if (-not $dt) { continue }
  if ($dt -lt $fromYearUtc) { continue }

  Update-Earliest -map $inviteByEmail -key $userMail -dt $dt
  $matchedRows++
}

Write-Host ("Invite rows matched (>= {0}) = {1}, distinct emails trouvés={2}" -f $PreciseFromYear, $matchedRows, $inviteByEmail.Count)

# ---- 3) Output
$out = foreach ($u in $users) {
  $em = ([string]$u.Email).Trim().ToLowerInvariant()
  $dt = $null
  if ($em -and $inviteByEmail.ContainsKey($em)) { $dt = $inviteByEmail[$em] }

  $inviteValue = if ($dt) { $dt.ToString("o") } else { "Avant $PreciseFromYear" }
  $inviteIsPrecise = [bool]$dt
  $inviteSource = if ($dt) { "AUDIT_LOG_USER_EMAIL" } else { "BUCKET_NOT_FOUND_SINCE_2025" }

  [pscustomobject]@{
    Email                  = $u.Email
    AccountId              = $u.AccountId
    InviteDateOrBefore2025 = $inviteValue
    InviteIsPrecise        = $inviteIsPrecise
    InviteSource           = $inviteSource
  }
}

$outCsv = Join-Path $ExportsDir ("UserSansApps_WithInviteDate_{0}.csv" -f $ts)
$out | Export-Csv -Path $outCsv -Delimiter $delimUsers -NoTypeInformation -Encoding UTF8
Write-Host "Export -> $outCsv"