[CmdletBinding()]
param(
  # Bridage org Jiradot
  [string] $AllowedOrgId = "51960987-c893-423d-8ec8-8682042ec47b",

  # Cutoff (YYYY-MM-DD) : inactifs AVANT cette date
  [string] $CutoffDate = "2026-04-01",

  # Domaines autorisés
  [string[]] $AllowedDomainsInput = @(
    "chorum.fr",
    "harmonie-mutuelle.fr",
    "harmonie-sante.fr",
    "mutex-exterieur.fr",
    "mutex.fr",
    "prestataire.harmonie-mutuelle.fr",
    "prestataire.sihm.fr"
  ),

  # Si présent : suspend les NoApps classés INACTIVE_BEFORE_CUTOFF (jamais actif exclu)
  [switch] $ApplySuspend,

  # Sécurité: si UNKNOWN>0, on bloque la suspension
  [switch] $BlockSuspendIfUnknown = $true,

  # Rate limit / réseau
  [int] $ThrottleMs = 400,
  [int] $MaxRetries = 10
)

try { Set-StrictMode -Off } catch {}

Add-Type -AssemblyName "System.Globalization" | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("NoApps-LastActive-Cutoff_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "ScriptDir=$ScriptDir"
Log "ExportsDir=$ExportsDir"
Log ("ApplySuspend={0} BlockSuspendIfUnknown={1}" -f $ApplySuspend, $BlockSuspendIfUnknown)

# -------------------- Network (TLS + proxy) --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
  Log "Proxy système initialisé (DefaultWebProxy)."
} catch {
  Log "Proxy init impossible: $($_.Exception.Message)" "WARN"
}

# -------------------- Helpers --------------------
function Normalize-Domain([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.Trim()
  if ($s.StartsWith("@")) { $s = $s.Substring(1) }
  $s = $s -replace '^https?://',''
  $s = $s.Split('/')[0]
  return $s.ToLowerInvariant()
}

function Get-EmailDomain([string]$email) {
  if ([string]::IsNullOrWhiteSpace($email)) { return $null }
  $email = $email.Trim()
  $at = $email.LastIndexOf("@")
  if ($at -lt 0) { return $null }
  return $email.Substring($at + 1).ToLowerInvariant()
}

function Domain-IsAllowed([string]$emailDomain, [string[]]$allowedDomainsNormalized) {
  if ([string]::IsNullOrWhiteSpace($emailDomain)) { return $false }
  foreach ($d in $allowedDomainsNormalized) {
    if ([string]::IsNullOrWhiteSpace($d)) { continue }
    if ($emailDomain -eq $d) { return $true }
    if ($emailDomain.EndsWith("." + $d)) { return $true }
  }
  return $false
}

# ISO / epoch (sec/ms)
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

# Max-Date null-safe
function Max-Date {
  param(
    [AllowNull()] $a,
    [AllowNull()] $b
  )

  $da = if ($a -is [datetime]) { $a } else { $null }
  $db = if ($b -is [datetime]) { $b } else { $null }

  if ($da -and $db) {
    if ($da -ge $db) { return $da } else { return $db }
  }
  if ($da) { return $da }
  if ($db) { return $db }
  return $null
}

# Cutoff parsing
try {
  $cutoffLocal = [datetime]::ParseExact($CutoffDate, "yyyy-MM-dd", $culture)
  $cutoffUtc = [DateTime]::SpecifyKind($cutoffLocal, [DateTimeKind]::Local).ToUniversalTime()
} catch {
  throw "CutoffDate invalide: $CutoffDate (attendu YYYY-MM-DD)"
}
Log ("CutoffDate={0} (UTC={1})" -f $CutoffDate, $cutoffUtc.ToString("o"))

$AllowedDomains = @($AllowedDomainsInput | ForEach-Object { Normalize-Domain $_ } | Where-Object { $_ } | Sort-Object -Unique)
Log ("AllowedDomains={0}" -f ([string]::Join(", ", $AllowedDomains)))

# -------------------- Load org-admin.xml --------------------
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) {
  throw "Fichier manquant: $orgCredFile (créé via Save-AtlassianOrgAdminKey.ps1)"
}

$data = Import-Clixml -Path $orgCredFile
$orgId = [string]$data.OrgId
$apiKeyPlain = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password

if ([string]::IsNullOrWhiteSpace($orgId) -or [string]::IsNullOrWhiteSpace($apiKeyPlain)) {
  throw "orgId / apiKey vides dans $orgCredFile"
}
if ($orgId -ne $AllowedOrgId) {
  throw "OrgId non autorisé. Attendu '$AllowedOrgId', trouvé '$orgId'."
}

$headers = @{ Authorization = "Bearer $apiKeyPlain"; Accept = "application/json" }
Log "OrgId validé JIRADOT = $orgId"

# -------------------- HTTP with retry --------------------
function Invoke-AdminRequest {
  param(
    [ValidateSet("GET","POST")][string]$Method,
    [string]$Url,
    $Body = $null
  )

  $attempt = 0
  while ($true) {
    $attempt++
    try {
      if ($Method -eq "POST") {
        $json = ($Body | ConvertTo-Json -Depth 20)
        $resp = Invoke-WebRequest -Method POST -Uri $Url -Headers $script:headers -UseBasicParsing `
          -ContentType "application/json" -Body $json -ErrorAction Stop
      } else {
        $resp = Invoke-WebRequest -Method GET -Uri $Url -Headers $script:headers -UseBasicParsing -ErrorAction Stop
      }

      if ([string]::IsNullOrWhiteSpace($resp.Content)) { return $null }
      return ($resp.Content | ConvertFrom-Json)
    } catch {
      $status = 0
      $retryAfter = $null
      try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
      try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch {}

      if ($attempt -gt $MaxRetries) {
        throw "$Method failed after $MaxRetries retries. status=$status url=$Url details=$($_.Exception.Message)"
      }

      $retryable = ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0)
      if ($retryable) {
        $sleepSec = 0
        if ($retryAfter) {
          $tmp = 0
          if ([int]::TryParse([string]$retryAfter, [ref]$tmp) -and $tmp -gt 0) { $sleepSec = $tmp }
        }
        if ($sleepSec -le 0) { $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6,$attempt))) }

        Log "Retry $Method status=$status in ${sleepSec}s (attempt $attempt/$MaxRetries)" "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }

      throw "$Method error status=$status url=$Url details=$($_.Exception.Message)"
    }
  }
}

function AdminGet([string]$Url) { Invoke-AdminRequest -Method "GET" -Url $Url }

# -------------------- API: list org users v1 --------------------
function Get-AllOrgUsersV1 {
  $users = New-Object System.Collections.Generic.List[object]
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/users"

  while ($true) {
    Log "GET $url"
    $resp = AdminGet $url
    foreach ($u in @($resp.data)) { $users.Add($u) | Out-Null }

    $next = $null
    try { $next = [string]$resp.links.next } catch {}
    if ([string]::IsNullOrWhiteSpace($next)) { break }

    if ($next -match "^https?://") { $url = $next } else { $url = "https://api.atlassian.com$next" }
    Start-Sleep -Milliseconds $ThrottleMs
  }

  return $users
}

# -------------------- API: suspend-access v1 --------------------
function Suspend-AccessV1 {
  param([string]$AccountId)
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$AccountId/suspend-access"
  return Invoke-AdminRequest -Method "POST" -Url $url -Body @{}
}

# ==================== MAIN ====================
$orgUsers = Get-AllOrgUsersV1
Log ("Org users fetched (v1/users) = {0}" -f $orgUsers.Count)

$domainUsers = New-Object System.Collections.Generic.List[object]
$appRows = New-Object System.Collections.Generic.List[object]

foreach ($u in $orgUsers) {
  $accId = [string]$u.account_id
  if ([string]::IsNullOrWhiteSpace($accId)) { continue }

  $email = [string]$u.email
  $dom = Get-EmailDomain $email
  if (-not (Domain-IsAllowed -emailDomain $dom -allowedDomainsNormalized $AllowedDomains)) { continue }

  $orgLastUtc = Parse-AtlassianDateUtc $u.last_active

  $pa = @()
  try { $pa = @($u.product_access) } catch { $pa = @() }

  $apps = @()
  $maxAppUtc = $null

  if ($pa.Count -gt 0) {
    foreach ($p in $pa) {
      $k = [string]$p.key
      if ($k) { $apps += $k }

      $pLastStr = $null
      $pLastUtc = $null

      $pLastStr = [string]$p.last_active_timestamp
      $pLastUtc = Parse-AtlassianDateUtc $pLastStr
      if (-not $pLastUtc) {
        $pLastStr = [string]$p.last_active
        $pLastUtc = Parse-AtlassianDateUtc $pLastStr
      }
      if (-not $pLastUtc) {
        $pLastStr = [string]$p.last_active_date
        $pLastUtc = Parse-AtlassianDateUtc $pLastStr
      }

      $maxAppUtc = Max-Date $maxAppUtc $pLastUtc

      $appRows.Add([pscustomobject]@{
        AccountId        = $accId
        Email            = $email
        EmailDomain      = $dom
        AppKey           = $k
        AppUrl           = [string]$p.url
        AppLastActive    = $pLastStr
        AppLastActiveUtc = $(if($pLastUtc){$pLastUtc.ToString("o")}else{$null})
      }) | Out-Null
    }
  } else {
    $appRows.Add([pscustomobject]@{
      AccountId        = $accId
      Email            = $email
      EmailDomain      = $dom
      AppKey           = $null
      AppUrl           = $null
      AppLastActive    = $null
      AppLastActiveUtc = $null
    }) | Out-Null
  }

  $effectiveUtc = Max-Date $maxAppUtc $orgLastUtc

  $domainUsers.Add([pscustomobject]@{
    AccountId              = $accId
    Email                  = $email
    EmailDomain            = $dom
    Name                   = [string]$u.name
    AccountStatus          = [string]$u.account_status
    AccessBillable         = [string]$u.access_billable

    ProductAccessCount     = $pa.Count
    Apps                   = (($apps | Sort-Object -Unique) -join ",")

    OrgLastActive          = [string]$u.last_active
    OrgLastActiveUtc       = $(if($orgLastUtc){$orgLastUtc.ToString("o")}else{$null})

    MaxAppLastActiveUtc    = $(if($maxAppUtc){$maxAppUtc.ToString("o")}else{$null})
    EffectiveLastActiveUtc = $(if($effectiveUtc){$effectiveUtc.ToString("o")}else{$null})
  }) | Out-Null
}

Log ("Users after domain filter = {0}" -f $domainUsers.Count)

# ---------- Exports FAST ----------
$domainCsv = Join-Path $ExportsDir ("DomainUsers_Domains_{0}.csv" -f $ts)
$domainUsers | Sort-Object EmailDomain, Email | Export-Csv -Path $domainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $domainCsv"

$appCsv = Join-Path $ExportsDir ("DomainUsers_AppLastActive_{0}.csv" -f $ts)
$appRows | Sort-Object EmailDomain, Email, AppKey | Export-Csv -Path $appCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $appCsv"

# ---------- NoApps + classification ----------
$noApps = @($domainUsers | Where-Object { $_.ProductAccessCount -eq 0 })
Log ("NoApps (FAST) = {0}" -f $noApps.Count)

foreach ($x in $noApps) {
  if ([string]::IsNullOrWhiteSpace($x.EffectiveLastActiveUtc)) {
    # "Jamais actif" -> traité à part
    $x | Add-Member -NotePropertyName "LastActiveClass" -NotePropertyValue "NEVER_ACTIVE" -Force
  } else {
    $dt = Parse-AtlassianDateUtc $x.EffectiveLastActiveUtc
    if (-not $dt) {
      $x | Add-Member -NotePropertyName "LastActiveClass" -NotePropertyValue "UNKNOWN" -Force
    } elseif ($dt -lt $cutoffUtc) {
      $x | Add-Member -NotePropertyName "LastActiveClass" -NotePropertyValue "INACTIVE_BEFORE_CUTOFF" -Force
    } else {
      $x | Add-Member -NotePropertyName "LastActiveClass" -NotePropertyValue "ACTIVE_AFTER_CUTOFF" -Force
    }
  }
}

$noAppsAllCsv = Join-Path $ExportsDir ("NoApps_All_{0}.csv" -f $ts)
$noApps | Sort-Object LastActiveClass, Email | Export-Csv -Path $noAppsAllCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $noAppsAllCsv"

$noAppsNever   = @($noApps | Where-Object { $_.LastActiveClass -eq "NEVER_ACTIVE" })
$noAppsBefore  = @($noApps | Where-Object { $_.LastActiveClass -eq "INACTIVE_BEFORE_CUTOFF" })
$noAppsAfter   = @($noApps | Where-Object { $_.LastActiveClass -eq "ACTIVE_AFTER_CUTOFF" })
$noAppsUnknown = @($noApps | Where-Object { $_.LastActiveClass -eq "UNKNOWN" })

$cut = $CutoffDate.Replace("-","")

$neverCsv = Join-Path $ExportsDir ("NoApps_NeverActive_{0}.csv" -f $ts)
$noAppsNever | Sort-Object Email | Export-Csv -Path $neverCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $neverCsv"

$beforeCsv = Join-Path $ExportsDir ("NoApps_InactiveBeforeCutoff_{0}_{1}.csv" -f $cut, $ts)
$noAppsBefore | Sort-Object EffectiveLastActiveUtc | Export-Csv -Path $beforeCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $beforeCsv"

$afterCsv = Join-Path $ExportsDir ("NoApps_ActiveAfterCutoff_{0}_{1}.csv" -f $cut, $ts)
$noAppsAfter | Sort-Object EffectiveLastActiveUtc | Export-Csv -Path $afterCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $afterCsv"

$unknownCsv = Join-Path $ExportsDir ("NoApps_UnknownLastActive_{0}.csv" -f $ts)
$noAppsUnknown | Sort-Object Email | Export-Csv -Path $unknownCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $unknownCsv"

Log ("NoApps classes: NEVER_ACTIVE={0} INACTIVE_BEFORE_CUTOFF={1} ACTIVE_AFTER_CUTOFF={2} UNKNOWN={3}" -f `
  $noAppsNever.Count, $noAppsBefore.Count, $noAppsAfter.Count, $noAppsUnknown.Count)

# Summary NoApps by domain
$sumNoAppsDomain =
  $noApps |
  Group-Object EmailDomain |
  ForEach-Object {
    $g = $_.Group
    [pscustomobject]@{
      EmailDomain             = $_.Name
      NoAppsTotal             = $g.Count
      NeverActive             = ($g | Where-Object { $_.LastActiveClass -eq "NEVER_ACTIVE" }).Count
      InactiveBeforeCutoff    = ($g | Where-Object { $_.LastActiveClass -eq "INACTIVE_BEFORE_CUTOFF" }).Count
      ActiveAfterCutoff       = ($g | Where-Object { $_.LastActiveClass -eq "ACTIVE_AFTER_CUTOFF" }).Count
      UnknownLastActive       = ($g | Where-Object { $_.LastActiveClass -eq "UNKNOWN" }).Count
    }
  } | Sort-Object NoAppsTotal -Descending

$sumNoAppsDomainCsv = Join-Path $ExportsDir ("Summary_NoApps_ByDomain_{0}.csv" -f $ts)
$sumNoAppsDomain | Export-Csv -Path $sumNoAppsDomainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $sumNoAppsDomainCsv"

# ---------- Apply suspend (guard-railed) ----------
if ($ApplySuspend) {
  if ($BlockSuspendIfUnknown -and $noAppsUnknown.Count -gt 0) {
    throw "SECURITE: suspension interdite car il reste $($noAppsUnknown.Count) NoApps avec last active UNKNOWN."
  }

  if ($noAppsBefore.Count -eq 0) {
    Log "Aucun NoApps INACTIVE_BEFORE_CUTOFF à suspendre." "OK"
  } else {
    Log "APPLY: suspension des NoApps INACTIVE_BEFORE_CUTOFF (count=$($noAppsBefore.Count)). (NEVER_ACTIVE exclus)" "WARN"

    $actions = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($u in $noAppsBefore) {
      $i++
      Write-Progress -Activity "Suspend NoApps INACTIVE_BEFORE_CUTOFF" -Status "$i / $($noAppsBefore.Count) : $($u.Email)" -PercentComplete ([int](100*$i/$noAppsBefore.Count))

      try {
        Suspend-AccessV1 -AccountId $u.AccountId | Out-Null
        $actions.Add([pscustomobject]@{ AccountId=$u.AccountId; Email=$u.Email; Action="SUSPEND_ACCESS"; Status="OK" }) | Out-Null
        Log "OK suspend-access $($u.Email)" "OK"
      } catch {
        $actions.Add([pscustomobject]@{ AccountId=$u.AccountId; Email=$u.Email; Action="SUSPEND_ACCESS"; Status="ERROR"; Details=$_.Exception.Message }) | Out-Null
        Log "ERROR suspend-access $($u.Email): $($_.Exception.Message)" "ERROR"
      }

      Start-Sleep -Milliseconds $ThrottleMs
    }

    $actionsCsv = Join-Path $ExportsDir ("Actions_Suspend_NoAppsInactiveBeforeCutoff_{0}.csv" -f $ts)
    $actions | Export-Csv -Path $actionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Log "Export -> $actionsCsv"
  }
}

Log "Terminé."