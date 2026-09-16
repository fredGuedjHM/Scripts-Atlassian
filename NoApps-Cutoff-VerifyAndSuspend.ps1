[CmdletBinding()]
param(
  # Bridage org Jiradot
  [string] $AllowedOrgId = "51960987-c893-423d-8ec8-8682042ec47b",

  # Cutoff: "ajoutés à l'org avant" cette date
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

  # Contrôle appels last-active-dates (uniquement pour NoApps manquants du cache)
  [int] $MaxLastActiveDatesCalls = 200,

  # Appliquer la suspension (uniquement NoApps + BEFORE_CUTOFF + zéro UNKNOWN)
  [switch] $ApplySuspend,

  # Debug: exporter un échantillon de payload last-active-dates (utile si encore des soucis de parsing)
  [int] $DebugLastActiveDatesSamples = 0,

  # Réseau / rate-limit
  [int] $ThrottleMs = 500,
  [int] $MaxRetries = 10
)

try { Set-StrictMode -Off } catch {}

Add-Type -AssemblyName "System.Globalization" | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
$CacheDir   = Join-Path $ScriptDir "cache"

foreach ($p in @($ExportsDir, $CacheDir)) {
  if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("NoApps-Cutoff-VerifyAndSuspend_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "ScriptDir=$ScriptDir"
Log "ExportsDir=$ExportsDir"
Log "CacheDir=$CacheDir"
Log ("ApplySuspend={0}" -f $ApplySuspend)
Log ("MaxLastActiveDatesCalls={0}" -f $MaxLastActiveDatesCalls)
Log ("DebugLastActiveDatesSamples={0}" -f $DebugLastActiveDatesSamples)

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

# CORRECTIF IMPORTANT: accepte ISO, epoch ms, epoch s, et aussi des nombres JSON (int64) directement
function Parse-AtlassianDateUtc($v) {
  if ($null -eq $v) { return $null }

  # valeur numérique JSON
  if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
    $n = [int64]$v
    if ($n -gt 1000000000000) { return [DateTimeOffset]::FromUnixTimeMilliseconds($n).UtcDateTime } # ms
    if ($n -gt 1000000000)    { return [DateTimeOffset]::FromUnixTimeSeconds($n).UtcDateTime }      # sec
    return $null
  }

  $s = [string]$v
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.Trim()

  # numérique en string
  if ($s -match '^\d+$') {
    $n = [int64]$s
    if ($n -gt 1000000000000) { return [DateTimeOffset]::FromUnixTimeMilliseconds($n).UtcDateTime }
    if ($n -gt 1000000000)    { return [DateTimeOffset]::FromUnixTimeSeconds($n).UtcDateTime }
    return $null
  }

  # ISO / RFC3339 / etc.
  try { return [DateTimeOffset]::Parse($s, $culture).UtcDateTime } catch { return $null }
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
  throw "OrgId non autorisé. Attendu '$AllowedOrgId' (JIRADOT), trouvé '$orgId'."
}

$headers = @{ Authorization = "Bearer $apiKeyPlain"; Accept = "application/json" }
Log "OrgId validé JIRADOT = $orgId"

# -------------------- HTTP with retry (GET/POST, 429/5xx/status=0) + throttle auto --------------------
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
        $json = ($Body | ConvertTo-Json -Depth 30)
        $resp = Invoke-WebRequest -Method POST -Uri $Url -Headers $script:headers -UseBasicParsing `
          -ContentType "application/json" -Body $json -ErrorAction Stop
      } else {
        $resp = Invoke-WebRequest -Method GET -Uri $Url -Headers $script:headers -UseBasicParsing -ErrorAction Stop
      }

      # Cooldown léger sur succès
      if ($script:ThrottleMs -gt 500) { $script:ThrottleMs = [Math]::Max(500, $script:ThrottleMs - 25) }

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
        if ($status -eq 429) {
          $script:ThrottleMs = [Math]::Min(5000, $script:ThrottleMs + 250)
          Log "Rate-limit: ThrottleMs augmenté à $script:ThrottleMs ms" "WARN"
        }

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
function AdminPost([string]$Url, $Body) { Invoke-AdminRequest -Method "POST" -Url $Url -Body $Body }

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

# -------------------- API: last-active-dates (preuve AddedToOrg) --------------------
function Get-LastActiveDatesV1 {
  param([string]$AccountId)
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$AccountId/last-active-dates"
  return AdminGet $url
}

# -------------------- API: suspend-access (NoApps only) --------------------
function Suspend-AccessV1 {
  param([string]$AccountId)
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$AccountId/suspend-access"
  return Invoke-AdminRequest -Method "POST" -Url $url -Body @{}
}

# -------------------- Cache addedToOrg (accountId -> ISO string) --------------------
$addedCachePath = Join-Path $CacheDir "addedToOrgCache.json"
$addedCache = @{}
if (Test-Path $addedCachePath) {
  try {
    $tmp = Get-Content -Raw -Path $addedCachePath -Encoding UTF8 | ConvertFrom-Json
    if ($tmp) {
      foreach ($p in $tmp.PSObject.Properties) { $addedCache[$p.Name] = [string]$p.Value }
      Log ("Cache loaded: {0} entries" -f $addedCache.Count)
    }
  } catch {
    Log "WARN cannot read cache $addedCachePath : $($_.Exception.Message)" "WARN"
  }
}

function Save-AddedCache {
  try {
    ($addedCache | ConvertTo-Json -Depth 5) | Set-Content -Path $addedCachePath -Encoding UTF8
    Log ("Cache saved: {0}" -f $addedCachePath)
  } catch {
    Log "WARN cannot write cache $addedCachePath : $($_.Exception.Message)" "WARN"
  }
}

# ==================== MAIN ====================
$orgUsers = Get-AllOrgUsersV1
Log ("Org users fetched (v1/users) = {0}" -f $orgUsers.Count)

# FAST: domain filter + compute NoApps + apps/appLastActive (from /users payload)
$domainUsers = New-Object System.Collections.Generic.List[object]
$appRows = New-Object System.Collections.Generic.List[object]

foreach ($u in $orgUsers) {
  $accId = [string]$u.account_id
  if ([string]::IsNullOrWhiteSpace($accId)) { continue }

  $email = [string]$u.email
  $dom = Get-EmailDomain $email
  if (-not (Domain-IsAllowed -emailDomain $dom -allowedDomainsNormalized $AllowedDomains)) { continue }

  $pa = @()
  try { $pa = @($u.product_access) } catch { $pa = @() }

  $apps = @()
  $maxAppUtc = $null

  if ($pa.Count -gt 0) {
    foreach ($p in $pa) {
      $k = [string]$p.key
      if ($k) { $apps += $k }

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

      if ($pLastUtc -and ($null -eq $maxAppUtc -or $pLastUtc -gt $maxAppUtc)) { $maxAppUtc = $pLastUtc }

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

  $domainUsers.Add([pscustomobject]@{
    AccountId          = $accId
    Email              = $email
    EmailDomain        = $dom
    Name               = [string]$u.name
    AccountStatus      = [string]$u.account_status
    AccessBillable     = [string]$u.access_billable
    OrgLastActive      = [string]$u.last_active
    ProductAccessCount = $pa.Count
    Apps               = (($apps | Sort-Object -Unique) -join ",")
    MaxAppLastActiveUtc= $(if($maxAppUtc){$maxAppUtc.ToString("o")}else{$null})

    AddedToOrg         = $null
    AddedToOrgUtc      = $null
    AddedToOrgSource   = $null
    CutoffClass        = $null
  }) | Out-Null
}

Log ("Users after domain filter = {0}" -f $domainUsers.Count)

# Export FAST datasets
$domainCsv = Join-Path $ExportsDir ("DomainUsers_Domains_{0}.csv" -f $ts)
$domainUsers | Sort-Object EmailDomain, Email | Export-Csv -Path $domainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $domainCsv"

$appCsv = Join-Path $ExportsDir ("DomainUsers_AppLastActive_{0}.csv" -f $ts)
$appRows | Sort-Object EmailDomain, Email, AppKey | Export-Csv -Path $appCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $appCsv"

# Identify NoApps
$noApps = @($domainUsers | Where-Object { $_.ProductAccessCount -eq 0 })
Log ("NoApps (FAST) = {0}" -f $noApps.Count)

$noAppsFastCsv = Join-Path $ExportsDir ("Candidates_NoApps_{0}.csv" -f $ts)
$noApps | Sort-Object Email | Export-Csv -Path $noAppsFastCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $noAppsFastCsv"

# Debug samples (optional)
if ($DebugLastActiveDatesSamples -gt 0) {
  $dbgOut = New-Object System.Collections.Generic.List[object]
  foreach ($u in ($noApps | Select-Object -First $DebugLastActiveDatesSamples)) {
    try {
      $lad = Get-LastActiveDatesV1 -AccountId $u.AccountId
      $dbgOut.Add([pscustomobject]@{
        Email = $u.Email
        AccountId = $u.AccountId
        RawAddedToOrgTimestamp = $lad.data.added_to_org_timestamp
        RawAddedToOrg          = $lad.data.added_to_org
      }) | Out-Null
    } catch {
      $dbgOut.Add([pscustomobject]@{
        Email = $u.Email
        AccountId = $u.AccountId
        RawAddedToOrgTimestamp = $null
        RawAddedToOrg = $null
        Error = $_.Exception.Message
      }) | Out-Null
    }
    Start-Sleep -Milliseconds $ThrottleMs
  }

  $dbgCsv = Join-Path $ExportsDir ("Debug_LastActiveDatesSamples_{0}.csv" -f $ts)
  $dbgOut | Export-Csv -Path $dbgCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
  Log "Export -> $dbgCsv"
}

# -------------------- Resolve AddedToOrg ONLY for NoApps (limited + cached) --------------------
$calls = 0
$updated = 0

foreach ($u in $noApps) {
  if ($calls -ge $MaxLastActiveDatesCalls) { break }

  $accId = $u.AccountId
  if ($addedCache.ContainsKey($accId) -and -not [string]::IsNullOrWhiteSpace($addedCache[$accId])) {
    continue
  }

  $calls++
  Write-Progress -Activity "Resolve added_to_org (NoApps only)" -Status "$calls / $MaxLastActiveDatesCalls : $($u.Email)" -PercentComplete ([int](100*$calls/$MaxLastActiveDatesCalls))

  try {
    $lad = Get-LastActiveDatesV1 -AccountId $accId

    $addedRaw = $null
    try { $addedRaw = $lad.data.added_to_org_timestamp } catch {}
    if ($null -eq $addedRaw -or ([string]$addedRaw).Trim() -eq "") {
      try { $addedRaw = $lad.data.added_to_org } catch {}
    }

    $addedUtc = Parse-AtlassianDateUtc $addedRaw
    if ($addedUtc) {
      # Cache en ISO pour fiabiliser tous les runs suivants
      $addedCache[$accId] = $addedUtc.ToString("o")
      $updated++
    } else {
      Log "WARN AddedToOrg non parsable for $($u.Email) raw=[$addedRaw]" "WARN"
    }
  } catch {
    Log "WARN last-active-dates failed for $($u.Email) ($accId): $($_.Exception.Message)" "WARN"
  }

  Start-Sleep -Milliseconds $ThrottleMs
}

Save-AddedCache
Log ("Resolve AddedToOrg: calls={0}, cache updated={1}" -f $calls, $updated)

# Apply cache to NoApps + classification cutoff
foreach ($u in $noApps) {
  $accId = $u.AccountId
  if ($addedCache.ContainsKey($accId)) {
    $addedIso = [string]$addedCache[$accId]
    $addedUtc = Parse-AtlassianDateUtc $addedIso

    $u.AddedToOrg = $addedIso
    $u.AddedToOrgUtc = $(if($addedUtc){$addedUtc.ToString("o")}else{$null})
    $u.AddedToOrgSource = "last-active-dates-cache"
  }

  if ([string]::IsNullOrWhiteSpace($u.AddedToOrgUtc)) {
    $u.CutoffClass = "UNKNOWN"
  } else {
    $dt = Parse-AtlassianDateUtc $u.AddedToOrgUtc
    if ($dt -and $dt -lt $cutoffUtc) { $u.CutoffClass = "BEFORE_CUTOFF" }
    elseif ($dt) { $u.CutoffClass = "AFTER_CUTOFF" }
    else { $u.CutoffClass = "UNKNOWN" }
  }
}

$noAppsResolvedCsv = Join-Path $ExportsDir ("NoApps_Resolved_{0}.csv" -f $ts)
$noApps | Sort-Object CutoffClass, Email | Export-Csv -Path $noAppsResolvedCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $noAppsResolvedCsv"

$noAppsBefore = @($noApps | Where-Object { $_.CutoffClass -eq "BEFORE_CUTOFF" })
$noAppsAfter  = @($noApps | Where-Object { $_.CutoffClass -eq "AFTER_CUTOFF" })
$noAppsUnknown= @($noApps | Where-Object { $_.CutoffClass -eq "UNKNOWN" })

$cut = $CutoffDate.Replace("-","")

$beforeCsv = Join-Path $ExportsDir ("NoApps_BeforeCutoff_{0}_{1}.csv" -f $cut, $ts)
$noAppsBefore | Sort-Object AddedToOrgUtc | Export-Csv -Path $beforeCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $beforeCsv"

$afterCsv = Join-Path $ExportsDir ("NoApps_AfterCutoff_{0}_{1}.csv" -f $cut, $ts)
$noAppsAfter | Sort-Object AddedToOrgUtc | Export-Csv -Path $afterCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $afterCsv"

$unknownCsv = Join-Path $ExportsDir ("NoApps_UnknownAddedToOrg_{0}.csv" -f $ts)
$noAppsUnknown | Sort-Object Email | Export-Csv -Path $unknownCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $unknownCsv"

Log ("NoApps: BEFORE={0} AFTER={1} UNKNOWN={2}" -f $noAppsBefore.Count, $noAppsAfter.Count, $noAppsUnknown.Count)

# Summary NoApps by domain (cutoff classes)
$sumNoAppsDomain =
  $noApps |
  Group-Object EmailDomain |
  ForEach-Object {
    $g = $_.Group
    [pscustomobject]@{
      EmailDomain       = $_.Name
      NoAppsTotal       = $g.Count
      BeforeCutoff      = ($g | Where-Object { $_.CutoffClass -eq "BEFORE_CUTOFF" }).Count
      AfterCutoff       = ($g | Where-Object { $_.CutoffClass -eq "AFTER_CUTOFF" }).Count
      UnknownAddedToOrg = ($g | Where-Object { $_.CutoffClass -eq "UNKNOWN" }).Count
    }
  } | Sort-Object NoAppsTotal -Descending

$sumNoAppsDomainCsv = Join-Path $ExportsDir ("Summary_NoApps_ByDomain_{0}.csv" -f $ts)
$sumNoAppsDomain | Export-Csv -Path $sumNoAppsDomainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $sumNoAppsDomainCsv"

# -------------------- Apply suspend (guard-railed) --------------------
if ($ApplySuspend) {
  if ($noAppsUnknown.Count -gt 0) {
    throw "SECURITE: suspension interdite car il reste $($noAppsUnknown.Count) NoApps avec AddedToOrg UNKNOWN. Relance avec MaxLastActiveDatesCalls plus grand."
  }

  if ($noAppsBefore.Count -eq 0) {
    Log "Aucun NoApps BEFORE_CUTOFF à suspendre." "OK"
  } else {
    Log "APPLY: suspension des NoApps BEFORE_CUTOFF (count=$($noAppsBefore.Count))" "WARN"

    $actions = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($u in $noAppsBefore) {
      $i++
      Write-Progress -Activity "Suspend NoApps BEFORE_CUTOFF" -Status "$i / $($noAppsBefore.Count) : $($u.Email)" -PercentComplete ([int](100*$i/$noAppsBefore.Count))

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

    $actionsCsv = Join-Path $ExportsDir ("Actions_Suspend_NoAppsBeforeCutoff_{0}.csv" -f $ts)
    $actions | Export-Csv -Path $actionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Log "Export -> $actionsCsv"
  }
}

Log "Terminé."