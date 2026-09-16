[CmdletBinding()]
param(
  # Bridage Jiradot org
  [string] $AllowedOrgId = "51960987-c893-423d-8ec8-8682042ec47b",

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

  # (Optionnel) Cutoff pour "ajouté à l'org avant"
  [string] $CutoffDate = "2026-04-01",

  # Par défaut: rapport FAST sans appels last-active-dates (recommandé)
  [switch] $FastReportOnly = $true,

  # Si activé: tente de récupérer AddedToOrg via last-active-dates (limité + cache),
  # puis exporte aussi un CSV "BeforeCutoff"
  [switch] $ResolveAddedToOrgViaLastActiveDates,

  # Nombre max d'appels last-active-dates par exécution (si ResolveAddedToOrgViaLastActiveDates)
  [int] $MaxLastActiveDatesCalls = 200,

  # Appliquer une suspension uniquement sur les "NoApps" (si tu l'actives, à manier prudemment)
  [switch] $ApplySuspendNoApps,

  # Rate-limit / réseau
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
$CacheDir   = Join-Path $ScriptDir "cache"

foreach ($p in @($ExportsDir, $CacheDir)) {
  if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("Suspend-OldUsers-WithNoApps_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "ScriptDir=$ScriptDir"
Log "ExportsDir=$ExportsDir"
Log "CacheDir=$CacheDir"
Log ("FastReportOnly={0} ResolveAddedToOrgViaLastActiveDates={1} MaxLastActiveDatesCalls={2}" -f $FastReportOnly, $ResolveAddedToOrgViaLastActiveDates, $MaxLastActiveDatesCalls)
Log ("ApplySuspendNoApps={0}" -f $ApplySuspendNoApps)

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

function TryParseUtc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  try { return [DateTimeOffset]::Parse($s, $culture).UtcDateTime } catch { return $null }
}

# Cutoff parsing (utilisé seulement si on a une date "added to org" fiable)
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

# -------------------- HTTP with retry (GET/POST, 429/5xx/status=0) + backoff + throttle auto --------------------
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

      # petit "cooldown" (succès)
      if ($script:ThrottleMs -gt 400) { $script:ThrottleMs = [Math]::Max(400, $script:ThrottleMs - 25) }

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
          # ralentissement durable si rate-limit
          $script:ThrottleMs = [Math]::Min(5000, $script:ThrottleMs + 200)
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

# -------------------- API: last-active-dates v1 (optionnel, très cher) --------------------
function Get-LastActiveDatesV1 {
  param([string]$AccountId)
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$AccountId/last-active-dates"
  return AdminGet $url
}

# -------------------- Cache addedToOrg by accountId (pour éviter de refaire les calls) --------------------
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

# -------------------- MAIN --------------------
$orgUsers = Get-AllOrgUsersV1
Log ("Org users fetched (v1/users) = {0}" -f $orgUsers.Count)

# Domain filter + build FAST data from /users payload
$domainUsers = New-Object System.Collections.Generic.List[object]
$appRowsFast = New-Object System.Collections.Generic.List[object]

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

      # Certains tenants: last_active / last_active_timestamp / last_active_date
      $pLastStr = [string]$p.last_active_timestamp
      $pLastUtc = TryParseUtc $pLastStr
      if (-not $pLastUtc) {
        $pLastStr = [string]$p.last_active
        $pLastUtc = TryParseUtc $pLastStr
      }
      if (-not $pLastUtc) {
        $pLastStr = [string]$p.last_active_date
        $pLastUtc = TryParseUtc $pLastStr
      }

      if ($pLastUtc -and ($null -eq $maxAppUtc -or $pLastUtc -gt $maxAppUtc)) { $maxAppUtc = $pLastUtc }

      $appRowsFast.Add([pscustomobject]@{
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
    # Pas d'app: on met une ligne vide pour audit
    $appRowsFast.Add([pscustomobject]@{
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

    # AddedToOrg = inconnu en FAST (sauf si l'API le fournit un jour)
    AddedToOrg         = $null
    AddedToOrgUtc      = $null
    AddedToOrgSource   = $null
  }) | Out-Null
}

Log ("Users after domain filter = {0}" -f $domainUsers.Count)

# -------------------- Export FAST CSVs --------------------
$domainCsv = Join-Path $ExportsDir ("DomainUsers_Domains_{0}.csv" -f $ts)
$domainUsers | Sort-Object EmailDomain, Email | Export-Csv -Path $domainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $domainCsv"

$appCsv = Join-Path $ExportsDir ("DomainUsers_AppLastActive_{0}.csv" -f $ts)
$appRowsFast | Sort-Object EmailDomain, Email, AppKey | Export-Csv -Path $appCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $appCsv"

# Candidates "NoApps" (FAST)
$noAppsFast = $domainUsers | Where-Object { $_.ProductAccessCount -eq 0 }
$noAppsCsv = Join-Path $ExportsDir ("Candidates_NoApps_{0}.csv" -f $ts)
$noAppsFast | Sort-Object Email | Export-Csv -Path $noAppsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $noAppsCsv"

# -------------------- Summaries (contrôle résultat) --------------------
# Summary by domain
$sumDomain = $domainUsers |
  Group-Object EmailDomain |
  ForEach-Object {
    $users = $_.Group
    [pscustomobject]@{
      EmailDomain = $_.Name
      UsersCount  = $users.Count
      NoAppsCount = ($users | Where-Object { $_.ProductAccessCount -eq 0 }).Count
      WithAppsCount = ($users | Where-Object { $_.ProductAccessCount -gt 0 }).Count
      MaxAppLastActiveUtc = ($users | Where-Object { $_.MaxAppLastActiveUtc } | Sort-Object MaxAppLastActiveUtc -Descending | Select-Object -First 1).MaxAppLastActiveUtc
    }
  } | Sort-Object UsersCount -Descending

$sumDomainCsv = Join-Path $ExportsDir ("Summary_ByDomain_{0}.csv" -f $ts)
$sumDomain | Export-Csv -Path $sumDomainCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $sumDomainCsv"

# Summary by app
$sumApp = $appRowsFast |
  Where-Object { $_.AppKey } |
  Group-Object AppKey |
  ForEach-Object {
    $rows = $_.Group
    [pscustomobject]@{
      AppKey = $_.Name
      UsersCount = ($rows | Select-Object -ExpandProperty AccountId -Unique).Count
      RowsCount  = $rows.Count
      MaxAppLastActiveUtc = ($rows | Where-Object { $_.AppLastActiveUtc } | Sort-Object AppLastActiveUtc -Descending | Select-Object -First 1).AppLastActiveUtc
    }
  } | Sort-Object UsersCount -Descending

$sumAppCsv = Join-Path $ExportsDir ("Summary_ByApp_{0}.csv" -f $ts)
$sumApp | Export-Csv -Path $sumAppCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $sumAppCsv"

# Summary by domain + app
$sumDomainApp = $appRowsFast |
  Where-Object { $_.AppKey } |
  Group-Object EmailDomain, AppKey |
  ForEach-Object {
    $rows = $_.Group
    $name = $_.Name -split ',\s*'
    [pscustomobject]@{
      EmailDomain = $name[0]
      AppKey      = $name[1]
      UsersCount  = ($rows | Select-Object -ExpandProperty AccountId -Unique).Count
      RowsCount   = $rows.Count
      MaxAppLastActiveUtc = ($rows | Where-Object { $_.AppLastActiveUtc } | Sort-Object AppLastActiveUtc -Descending | Select-Object -First 1).AppLastActiveUtc
    }
  } | Sort-Object EmailDomain, AppKey

$sumDomainAppCsv = Join-Path $ExportsDir ("Summary_ByDomainAndApp_{0}.csv" -f $ts)
$sumDomainApp | Export-Csv -Path $sumDomainAppCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $sumDomainAppCsv"

# -------------------- Optional: resolve AddedToOrg via last-active-dates (LIMITED + cached) --------------------
if (-not $ResolveAddedToOrgViaLastActiveDates) {
  Log "ResolveAddedToOrgViaLastActiveDates désactivé. (FAST terminé)" "OK"
  if ($FastReportOnly) { Log "Terminé."; return }
}

if ($ResolveAddedToOrgViaLastActiveDates) {
  Log "ResolveAddedToOrgViaLastActiveDates activé (limité + cache)." "WARN"

  $calls = 0
  $updated = 0

  foreach ($du in $domainUsers) {
    if ($calls -ge $MaxLastActiveDatesCalls) { break }

    $accId = $du.AccountId
    if ($addedCache.ContainsKey($accId) -and -not [string]::IsNullOrWhiteSpace($addedCache[$accId])) {
      continue
    }

    $calls++
    Write-Progress -Activity "Resolve added_to_org via last-active-dates" -Status "$calls / $MaxLastActiveDatesCalls : $($du.Email)" -PercentComplete ([int](100*$calls/$MaxLastActiveDatesCalls))

    try {
      $lad = Get-LastActiveDatesV1 -AccountId $accId
      $addedStr = $null
      try { $addedStr = [string]$lad.data.added_to_org_timestamp } catch {}
      if ([string]::IsNullOrWhiteSpace($addedStr)) {
        try { $addedStr = [string]$lad.data.added_to_org } catch {}
      }

      if (-not [string]::IsNullOrWhiteSpace($addedStr)) {
        $addedCache[$accId] = $addedStr
        $updated++
      }
    } catch {
      Log "WARN last-active-dates failed for $($du.Email) ($accId): $($_.Exception.Message)" "WARN"
    }

    Start-Sleep -Milliseconds $ThrottleMs
  }

  Save-AddedCache
  Log ("last-active-dates calls={0}, cache updated={1}" -f $calls, $updated)
}

# Apply cache back to domainUsers and export BeforeCutoff if possible
foreach ($du in $domainUsers) {
  $accId = $du.AccountId
  if ($addedCache.ContainsKey($accId)) {
    $addedStr = [string]$addedCache[$accId]
    $addedUtc = TryParseUtc $addedStr
    $du.AddedToOrg = $addedStr
    $du.AddedToOrgUtc = $(if($addedUtc){$addedUtc.ToString("o")}else{$null})
    $du.AddedToOrgSource = "last-active-dates-cache"
  }
}

$knownAdded = $domainUsers | Where-Object { $_.AddedToOrgUtc }
$unknownAdded = $domainUsers | Where-Object { -not $_.AddedToOrgUtc }
Log ("AddedToOrg known={0} unknown={1}" -f $knownAdded.Count, $unknownAdded.Count)

$beforeCutoff = $knownAdded | Where-Object { (TryParseUtc $_.AddedToOrgUtc) -lt $cutoffUtc }
$beforeCutoffCsv = Join-Path $ExportsDir ("DomainUsers_BeforeCutoff_{0}_{1}.csv" -f $CutoffDate.Replace("-",""), $ts)
$beforeCutoff | Sort-Object AddedToOrgUtc | Export-Csv -Path $beforeCutoffCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $beforeCutoffCsv"

$unknownCsv = Join-Path $ExportsDir ("DomainUsers_UnknownAddedToOrg_{0}.csv" -f $ts)
$unknownAdded | Sort-Object Email | Export-Csv -Path $unknownCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export -> $unknownCsv"

# Optional: suspend no-apps (FAST definition: ProductAccessCount=0), uniquement si demandé
if ($ApplySuspendNoApps) {
  Log "APPLY suspend-access sur NoApps (prudence)..." "WARN"
  $actions = New-Object System.Collections.Generic.List[object]

  foreach ($c in $noAppsFast) {
    $accId = $c.AccountId
    $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/directory/users/$accId/suspend-access"
    try {
      AdminPost -Url $url -Body @{} | Out-Null
      $actions.Add([pscustomobject]@{ AccountId=$accId; Email=$c.Email; Action="SUSPEND_ACCESS"; Status="OK" }) | Out-Null
      Log "OK suspend-access $($c.Email)" "OK"
    } catch {
      $actions.Add([pscustomobject]@{ AccountId=$accId; Email=$c.Email; Action="SUSPEND_ACCESS"; Status="ERROR"; Details=$_.Exception.Message }) | Out-Null
      Log "ERROR suspend-access $($c.Email): $($_.Exception.Message)" "ERROR"
    }
    Start-Sleep -Milliseconds $ThrottleMs
  }

  $actionsCsv = Join-Path $ExportsDir ("Actions_SuspendNoApps_{0}.csv" -f $ts)
  $actions | Export-Csv -Path $actionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
  Log "Export -> $actionsCsv"
}

Log "Terminé."