[CmdletBinding()]
param(
  [int] $PreciseFromYear = 2025,
  [int] $ThrottleMs = 300,
  [int] $MaxRetries = 10
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Globalization | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# -------------------- Paths --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
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

function Parse-AtlassianDateUtc($v) {
  if ($null -eq $v) { return $null }

  if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
    $n = [int64]$v
    if ($n -gt 1000000000000) { return [DateTimeOffset]::FromUnixTimeMilliseconds($n).UtcDateTime }
    if ($n -gt 1000000000)    { return [DateTimeOffset]::FromUnixTimeSeconds($n).UtcDateTime }
    return $null
  }

  $s = [string]$v
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.Trim()

  if ($s -match '^\d+$') {
    $n = [int64]$s
    if ($n -gt 1000000000000) { return [DateTimeOffset]::FromUnixTimeMilliseconds($n).UtcDateTime }
    if ($n -gt 1000000000)    { return [DateTimeOffset]::FromUnixTimeSeconds($n).UtcDateTime }
    return $null
  }

  try { return [DateTimeOffset]::Parse($s, $culture).UtcDateTime } catch { return $null }
}

function Import-CsvAuto([string]$path) {
  $h = Get-Content -Path $path -TotalCount 1 -Encoding UTF8
  $delim = ','
  if ($h -match ';') { $delim = ';' }
  return @{ delim = $delim; rows = (Import-Csv -Path $path -Delimiter $delim) }
}

# -------------------- Load credentials --------------------
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }

$data = Import-Clixml -Path $orgCredFile
$orgId = [string]$data.OrgId
$apiKeyPlain = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password
if ([string]::IsNullOrWhiteSpace($orgId) -or [string]::IsNullOrWhiteSpace($apiKeyPlain)) {
  throw "orgId / apiKey vides dans $orgCredFile"
}
$headers = @{ Authorization = "Bearer $apiKeyPlain"; Accept = "application/json" }

# -------------------- HTTP GET with retry --------------------
function AdminGet([string]$Url) {
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $resp = Invoke-WebRequest -Method GET -Uri $Url -Headers $script:headers -UseBasicParsing -ErrorAction Stop
      if ([string]::IsNullOrWhiteSpace($resp.Content)) { return $null }
      return ($resp.Content | ConvertFrom-Json)
    } catch {
      $status = 0
      $retryAfter = $null
      try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
      try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch {}

      if ($attempt -gt $MaxRetries) {
        throw "GET failed after $MaxRetries retries. status=$status url=$Url details=$($_.Exception.Message)"
      }

      $retryable = ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0)
      if ($retryable) {
        $sleepSec = 0
        if ($retryAfter) {
          $tmp = 0
          if ([int]::TryParse([string]$retryAfter, [ref]$tmp) -and $tmp -gt 0) { $sleepSec = $tmp }
        }
        if ($sleepSec -le 0) { $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6,$attempt))) }
        Start-Sleep -Seconds $sleepSec
        continue
      }

      throw "GET error status=$status url=$Url details=$($_.Exception.Message)"
    }
  }
}

# -------------------- Select input CSV --------------------
$usersCsv = Pick-File "Sélectionner UsersSansApps.csv (avec AccountId)"
if ([string]::IsNullOrWhiteSpace($usersCsv) -or -not (Test-Path $usersCsv)) { throw "CSV introuvable." }

$pack = Import-CsvAuto $usersCsv
$delim = $pack.delim
$users = $pack.rows
if (-not $users -or $users.Count -eq 0) { throw "CSV vide: $usersCsv" }

# -------------------- Process --------------------
$fromYearUtc = [datetime]::SpecifyKind([datetime]::ParseExact("$PreciseFromYear-01-01","yyyy-MM-dd",$culture), [DateTimeKind]::Utc)

$out = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($u in $users) {
  $i++
  $accId = ([string]$u.AccountId).Trim()
  $email = [string]$u.Email

  Write-Progress -Activity "Fetch added_to_org (last-active-dates)" -Status "$i/$($users.Count) $email" -PercentComplete ([int](100*$i/$($users.Count)))

  $addedUtc = $null
  $addedStr = $null
  $source = "NONE"

  if (-not [string]::IsNullOrWhiteSpace($accId)) {
    $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$accId/last-active-dates"
    try {
      $resp = AdminGet $url
      $d = $resp.data

      # Priorité au timestamp si présent
      $addedUtc = Parse-AtlassianDateUtc $d.added_to_org_timestamp
      if (-not $addedUtc) { $addedUtc = Parse-AtlassianDateUtc $d.added_to_org }

      $addedStr = $d.added_to_org
      $source = "API_last-active-dates"
    } catch {
      $source = "API_ERROR"
    }
  }

  $inviteValue = $null
  $inviteIsPrecise = $false

  if ($addedUtc -and $addedUtc -ge $fromYearUtc) {
    $inviteValue = $addedUtc.ToString("o")
    $inviteIsPrecise = $true
  } elseif ($addedUtc) {
    $inviteValue = "Avant $PreciseFromYear"
    $inviteIsPrecise = $false
  } else {
    $inviteValue = "INCONNU"
    $inviteIsPrecise = $false
  }

  $out.Add([pscustomobject]@{
    Email = $email
    AccountId = $accId
    AddedToOrgRaw = $addedStr
    AddedToOrgUtc = $(if($addedUtc){$addedUtc.ToString("o")}else{$null})
    InviteDateOrBefore2025 = $inviteValue
    InviteIsPrecise = $inviteIsPrecise
    Source = $source
  }) | Out-Null

  Start-Sleep -Milliseconds $ThrottleMs
}

$outCsv = Join-Path $ExportsDir ("UserSansApps_WithAddedToOrg_{0}.csv" -f $ts)
$out | Export-Csv -Path $outCsv -Delimiter $delim -NoTypeInformation -Encoding UTF8
Write-Host "Export -> $outCsv"