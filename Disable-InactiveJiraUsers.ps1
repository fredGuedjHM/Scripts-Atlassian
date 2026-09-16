<#
Disable-InactiveJiraUsers.ps1

Améliorations :
- Lit l'URL Jira + le PSCredential depuis .\secrets\jira-jiradot.cred.xml
- Demande la date de référence au lancement (prompt), au lieu de forcer le 1er avril

Objectif :
- Identifier les utilisateurs qui ont accès à Jira (via groupes)
- Vérifier s'ils ont un événement de type "login" dans l'audit depuis une date donnée
- Optionnellement, retirer leur accès Jira (suppression des groupes) en mode -Apply

Sorties dans .\exports :
- ActiveUsers_*.csv
- AuditLoginHits_*.csv
- InactiveUsers_*.csv
- ActionsTaken_*.csv
#>

[CmdletBinding()]
param(
  # Si vide, prendra l'URL depuis le fichier de credentials
  [string] $JiraBaseUrl = "",

  # Si vide, sera demandé via Read-Host (défaut = aujourd'hui - 30 jours)
  [string] $SinceDate = "",

  # Groupes d'accès Jira
  [string[]] $AccessGroups = @("jira-software-users"),

  # Cible Credential Manager en fallback (optionnel)
  [string] $CredManTarget = "JIRA_JIRADOT",

  # Appliquer réellement les suppressions de groupes
  [switch] $Apply,

  # Throttle & retry
  [int] $ThrottleMs = 300,
  [int] $MaxRetries = 8
)

try { Set-StrictMode -Off } catch {}

Add-Type -AssemblyName "System.Globalization" | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

# -------------------- Paths & log --------------------
$ScriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("Disable-InactiveJiraUsers_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "ScriptDir=$ScriptDir"
Log "ExportsDir=$ExportsDir"

# -------------------- Network prereqs --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
  Log "Proxy système initialisé (DefaultWebProxy)."
} catch {
  Log "Proxy non initialisé : $($_.Exception.Message)" "WARN"
}

# -------------------- Credentials + URL loader --------------------
function Get-JiraConnection {
  param([string]$CredManTarget)

  $scriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
  $clixmlPath = Join-Path $scriptDir "secrets\jira-jiradot.cred.xml"

  # 1) CLIXML : on s'attend à y trouver un objet { JiraBaseUrl, Credential }
  if (Test-Path $clixmlPath) {
    try {
      $data = Import-Clixml -Path $clixmlPath
      if ($data -and $data.JiraBaseUrl -and $data.Credential) {
        Log "URL + credentials chargés depuis $clixmlPath"
        return $data
      }
    } catch {
      Log "Erreur Import-Clixml ($clixmlPath) : $($_.Exception.Message)" "WARN"
    }
  } else {
    Log "Fichier de credentials non trouvé : $clixmlPath" "WARN"
  }

  # 2) Fallback : Credential Manager + URL demandée
  $cred = $null
  try {
    if (Get-Module -ListAvailable -Name CredentialManager) {
      Import-Module CredentialManager -ErrorAction Stop | Out-Null
      $c = Get-StoredCredential -Target $CredManTarget
      if ($c -and $c.UserName -and $c.Password) {
        $sec = ConvertTo-SecureString $c.Password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($c.UserName, $sec)
        Log "Credentials chargés depuis Windows Credential Manager (Target='$CredManTarget')."
      }
    }
  } catch {
    Log "Erreur CredentialManager : $($_.Exception.Message)" "WARN"
  }

  if (-not $cred) {
    throw "Impossible de charger les credentials Jira. Crée d'abord le fichier via Save-JiraCredential.ps1."
  }

  $jiraUrl = Read-Host -Prompt "Jira base URL (ex: https://jiradot.atlassian.net)"
  $jiraUrl = $jiraUrl.TrimEnd("/")

  return [pscustomobject]@{
    JiraBaseUrl = $jiraUrl
    Credential  = $cred
  }
}

function New-BasicAuthHeader {
  param([pscredential]$Cred)
  $plain = $Cred.UserName + ":" + ($Cred.GetNetworkCredential().Password)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($plain)
  $b64   = [Convert]::ToBase64String($bytes)
  return @{ Authorization = "Basic $b64"; Accept = "application/json" }
}

# -------------------- HTTP helper with retry --------------------
function Invoke-JiraRest {
  param(
    [string]$Method,
    [string]$Url,
    [hashtable]$Headers,
    $Body = $null
  )

  $attempt = 0
  while ($true) {
    $attempt++
    try {
      if ($Body -ne $null) {
        return Invoke-RestMethod -Method $Method -Uri $Url -Headers $Headers `
          -Body ($Body | ConvertTo-Json -Depth 20) -ContentType "application/json" -ErrorAction Stop
      } else {
        return Invoke-RestMethod -Method $Method -Uri $Url -Headers $Headers -ErrorAction Stop
      }
    } catch {
      $status = $null
      $retryAfter = $null
      try { $status = [int]$_.Exception.Response.StatusCode } catch {}
      try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch {}

      $details = $_.Exception.Message
      try {
        $stream = $_.Exception.Response.GetResponseStream()
        if ($stream) {
          $sr = New-Object System.IO.StreamReader($stream)
          $body = $sr.ReadToEnd()
          if ($body) { $details = "$details | $body" }
        }
      } catch {}

      if ($attempt -gt $MaxRetries) {
        throw "Echec après $MaxRetries tentatives. status=$status url=$Url details=$details"
      }

      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599)) {
        $sleepSec = 0
        if ($retryAfter) {
          $tmp = 0
          if ([int]::TryParse([string]$retryAfter, [ref]$tmp) -and $tmp -gt 0) { $sleepSec = $tmp }
        }
        if ($sleepSec -le 0) { $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(6,$attempt))) }
        Log "Retry status=$status in ${sleepSec}s (attempt $attempt/$MaxRetries) url=$Url" "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }

      throw "Erreur non-retryable status=$status url=$Url details=$details"
    }
  }
}

# -------------------- Charger URL + creds --------------------
$conn = Get-JiraConnection -CredManTarget $CredManTarget

if ([string]::IsNullOrWhiteSpace($JiraBaseUrl)) {
  $JiraBaseUrl = $conn.JiraBaseUrl
}
$JiraBaseUrl = $JiraBaseUrl.TrimEnd("/")

$cred    = $conn.Credential
$headers = New-BasicAuthHeader -Cred $cred

# -------------------- SinceDate : demande si absent --------------------
if ([string]::IsNullOrWhiteSpace($SinceDate)) {
  $defaultSince = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
  $input = Read-Host -Prompt "Date de référence (YYYY-MM-DD, défaut=$defaultSince)"
  if ([string]::IsNullOrWhiteSpace($input)) {
    $SinceDate = $defaultSince
  } else {
    $SinceDate = $input
  }
}

# SinceDate -> UTC + format ISO attendu par Jira
try {
  $sinceLocal = [datetime]::ParseExact($SinceDate, "yyyy-MM-dd", $culture)
  $sinceUtc = [DateTime]::SpecifyKind($sinceLocal, [DateTimeKind]::Local).ToUniversalTime()
} catch {
  throw "SinceDate invalide. Format attendu: YYYY-MM-DD. Reçu: $SinceDate"
}
$fromIso = $sinceUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", $culture)

Log "JiraBaseUrl=$JiraBaseUrl"
Log "SinceDate=$SinceDate (sinceUtc=$($sinceUtc.ToString('o'))) fromIso=$fromIso"
Log "Groups=$([string]::Join(',', $AccessGroups))"
Log "Mode=$(if($Apply){'APPLY (modifie les groupes)'}else{'DRY-RUN (aucune modification)'})"

# -------------------- 1) Récupérer les membres des groupes d'accès --------------------
function Get-GroupMembers {
  param([string]$GroupName)

  $members = New-Object System.Collections.Generic.List[object]
  $startAt = 0
  $maxResults = 50

  while ($true) {
    $url = "$JiraBaseUrl/rest/api/3/group/member?groupname=$([uri]::EscapeDataString($GroupName))&startAt=$startAt&maxResults=$maxResults"
    $resp = Invoke-JiraRest -Method GET -Url $url -Headers $headers

    foreach ($u in $resp.values) { $members.Add($u) | Out-Null }

    $isLast = $false
    try {
      if ($resp.isLast -ne $null) { $isLast = [bool]$resp.isLast }
      elseif ($resp.total -ne $null) { $isLast = ($startAt + $maxResults) -ge [int]$resp.total }
      else { $isLast = ($resp.values.Count -lt $maxResults) }
    } catch {
      $isLast = ($resp.values.Count -lt $maxResults)
    }

    if ($isLast) { break }
    $startAt += $maxResults
    Start-Sleep -Milliseconds $ThrottleMs
  }

  return $members
}

$activeUsersMap = @{}
foreach ($g in $AccessGroups) {
  Log "Lecture membres du groupe: $g"
  $m = Get-GroupMembers -GroupName $g
  Log "  -> $($m.Count) membres"

  foreach ($u in $m) {
    if (-not $u.accountId) { continue }
    if (-not $activeUsersMap.ContainsKey($u.accountId)) {
      $activeUsersMap[$u.accountId] = [pscustomobject]@{
        AccountId    = $u.accountId
        DisplayName  = $u.displayName
        EmailAddress = $u.emailAddress
        Groups       = New-Object System.Collections.Generic.List[string]
      }
    }
    $activeUsersMap[$u.accountId].Groups.Add($g) | Out-Null
  }
}

$activeUsers = $activeUsersMap.Values
$activeCsv = Join-Path $ExportsDir ("ActiveUsers_{0}.csv" -f $ts)
$activeUsers |
  Select-Object AccountId, DisplayName, EmailAddress, @{n="Groups";e={$_.Groups -join ","}} |
  Export-Csv -Path $activeCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export ActiveUsers -> $activeCsv"

# -------------------- 2) Lire l'audit Jira depuis SinceDate --------------------
$offset = 0
$limit  = 1000

$loginHits = New-Object System.Collections.Generic.List[object]
$lastSeen  = @{} # accountId -> DateTime

Log "Lecture audit Jira depuis $SinceDate (fromIso=$fromIso)"

while ($true) {
  $fromEncoded = [uri]::EscapeDataString($fromIso)
  $url = "$JiraBaseUrl/rest/api/3/auditing/record?offset=$offset&limit=$limit&from=$fromEncoded"

  $resp = Invoke-JiraRest -Method GET -Url $url -Headers $headers

  $records = $resp.records
  if (-not $records -or $records.Count -eq 0) { break }

  foreach ($r in $records) {
    $created = $null
    try { $created = [DateTimeOffset]::Parse([string]$r.created).UtcDateTime } catch {}

    $summary = [string]$r.summary
    $category = [string]$r.category
    $eventSource = [string]$r.eventSource

    $authorAccountId = $null
    try { $authorAccountId = [string]$r.authorAccountId } catch {}
    if ([string]::IsNullOrWhiteSpace($authorAccountId)) {
      try { $authorAccountId = [string]$r.authorKey } catch {}
    }

    # Heuristique "login"
    $isLogin = $false
    $blob = ($summary + " " + $category + " " + $eventSource).ToLowerInvariant()
    if ($blob -match "login" -or $blob -match "logged" -or $blob -match "log in") { $isLogin = $true }

    if ($isLogin) {
      $loginHits.Add([pscustomobject]@{
        CreatedUtc      = $(if($created){$created.ToString("o")}else{$null})
        Summary         = $summary
        Category        = $category
        EventSource     = $eventSource
        AuthorAccountId = $authorAccountId
      }) | Out-Null

      if ($authorAccountId -and $created) {
        if (-not $lastSeen.ContainsKey($authorAccountId) -or $created -gt $lastSeen[$authorAccountId]) {
          $lastSeen[$authorAccountId] = $created
        }
      }
    }
  }

  $offset += $limit
  Start-Sleep -Milliseconds $ThrottleMs
  if ($records.Count -lt $limit) { break }
}

$auditCsv = Join-Path $ExportsDir ("AuditLoginHits_{0}.csv" -f $ts)
$loginHits | Export-Csv -Path $auditCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export AuditLoginHits -> $auditCsv (hits=$($loginHits.Count))"

# -------------------- 3) Construire la liste des inactifs --------------------
$inactive = New-Object System.Collections.Generic.List[object]
foreach ($u in $activeUsers) {
  $seen = $null
  if ($lastSeen.ContainsKey($u.AccountId)) { $seen = $lastSeen[$u.AccountId] }

  if ($null -eq $seen) {
    $inactive.Add([pscustomobject]@{
      AccountId    = $u.AccountId
      DisplayName  = $u.DisplayName
      EmailAddress = $u.EmailAddress
      Groups       = ($u.Groups -join ",")
      LastSeenUtc  = $null
      Reason       = "No login audit hit since $SinceDate"
    }) | Out-Null
  }
}

$inactiveCsv = Join-Path $ExportsDir ("InactiveUsers_{0}.csv" -f $ts)
$inactive | Export-Csv -Path $inactiveCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export InactiveUsers -> $inactiveCsv (inactive=$($inactive.Count))"

# -------------------- 4) Apply (optionnel) --------------------
$actions = New-Object System.Collections.Generic.List[object]

if (-not $Apply) {
  Log "DRY-RUN: aucune suppression effectuée. Relance avec -Apply pour appliquer." "OK"
} else {
  if ($inactive.Count -eq 0) {
    Log "Aucun utilisateur inactif à traiter." "OK"
  } else {
    Log "APPLY: suppression des utilisateurs inactifs des groupes d'accès." "WARN"

    foreach ($row in $inactive) {
      foreach ($g in ($row.Groups -split ",")) {
        $g2 = $g.Trim()
        if ([string]::IsNullOrWhiteSpace($g2)) { continue }

        $url = "$JiraBaseUrl/rest/api/3/group/user?groupname=$([uri]::EscapeDataString($g2))&accountId=$([uri]::EscapeDataString($row.AccountId))"
        try {
          Invoke-JiraRest -Method DELETE -Url $url -Headers $headers | Out-Null
          $actions.Add([pscustomobject]@{
            AccountId = $row.AccountId
            Email     = $row.EmailAddress
            Group     = $g2
            Action    = "REMOVED_FROM_GROUP"
            Status    = "OK"
          }) | Out-Null
          Log "OK removed $($row.EmailAddress) from $g2" "OK"
        } catch {
          $actions.Add([pscustomobject]@{
            AccountId = $row.AccountId
            Email     = $row.EmailAddress
            Group     = $g2
            Action    = "REMOVED_FROM_GROUP"
            Status    = "ERROR"
          }) | Out-Null
          Log "ERROR removing $($row.EmailAddress) from $g2 : $($_.Exception.Message)" "ERROR"
        }

        Start-Sleep -Milliseconds $ThrottleMs
      }
    }
  }
}

$actionsCsv = Join-Path $ExportsDir ("ActionsTaken_{0}.csv" -f $ts)
$actions | Export-Csv -Path $actionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Log "Export ActionsTaken -> $actionsCsv"

Log "Terminé."