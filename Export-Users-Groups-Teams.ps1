<#
.SYNOPSIS
  Export-Users-Groups-Teams.ps1 v1.4
  Pour chaque utilisateur Jira actif (domaines autorises) :
  groupes Jira, equipes Tempo, derniere equipe active, licence Tempo.

.NOTES
  Auteur  : Frederic GUEDJ
  Version : 1.4 — Aout 2026
  Fix     : Source etape 2 -> /group/member (maxResults=50, startAt classique)
             Comptes sans email conserves (email vide, non exclus)
             Show-Progress plafonnee a 100
             Filtre domaine applique uniquement si email non vide
#>

[CmdletBinding()]
param(
  [string] $SiteOnly        = "Jiradot",
  [int]    $MaxRetries      = 5,
  [int]    $TempoTimeoutSec = 30
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts      = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("UsersGroupsTeams_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("UsersGroupsTeams_{0}.csv"  -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Format-DateFr([string]$isoDate) {
  if (-not $isoDate) { return "" }
  try {
    $dt = [DateTimeOffset]::Parse($isoDate)
    return $dt.ToLocalTime().ToString("dd/MM/yyyy")
  } catch { return $isoDate }
}

function Show-Progress([string]$Activity, [string]$Status, [int]$Current, [int]$Total) {
  $pct = if ($Total -gt 0) { [Math]::Min(100, [int](100 * $Current / $Total)) } else { 0 }
  Write-Progress -Activity $Activity -Status $Status -PercentComplete $pct
}

# ============================================================
# RESEAU
# ============================================================

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# ============================================================
# HTTP HELPER
# ============================================================

function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [hashtable]$Headers,
        [string]$Body = $null, [int]$TimeoutSec = 30)
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{
        Method          = $Method
        Uri             = $Url
        Headers         = $Headers
        UseBasicParsing = $true
        ErrorAction     = "Stop"
        TimeoutSec      = $TimeoutSec
      }
      if ($Body) {
        $params["ContentType"] = "application/json; charset=utf-8"
        $params["Body"]        = [System.Text.Encoding]::UTF8.GetBytes($Body)
      }
      $resp        = Invoke-WebRequest @params
      $contentUtf8 = $resp.Content
      try {
        $stream = $resp.RawContentStream
        if ($stream -and $stream.CanSeek) {
          $stream.Position = 0
          $reader          = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
          $contentUtf8     = $reader.ReadToEnd(); $reader.Close()
        }
      } catch {
        try {
          $isoBytes    = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($resp.Content)
          $contentUtf8 = [System.Text.Encoding]::UTF8.GetString($isoBytes)
        } catch {}
      }
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$contentUtf8 }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status  = [int]$_.Exception.Response.StatusCode
        $rd      = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $rd.ReadToEnd(); $rd.Close()
      } catch {}
      if ($attempt -gt $MaxRetries) {
        return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
      }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(5, $attempt)))
        Log ("  Retry {0} status={1} in {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
        Start-Sleep -Seconds $sleepSec; continue
      }
      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# ============================================================
# CREDENTIALS
# ============================================================

function Import-SiteCredentials([string]$FilePath, [string]$SiteName) {
  if (-not (Test-Path $FilePath)) {
    [System.Windows.Forms.MessageBox]::Show(
      ("Credentials pour {0} manquants." -f $SiteName), "Credentials $SiteName",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl       = Read-Host "URL du site (ex: jiradot.atlassian.net)"
    $inputUrl       = $inputUrl -replace "^https?://","" -replace "/.*$","" -replace "/$",""
    $adminEmail     = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $FilePath
  }
  $data  = Import-Clixml -Path $FilePath
  $url   = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization="Basic $auth"; Accept="application/json" }
    Name    = $SiteName
  }
}

function Import-TempoCredentials([string]$FilePath) {
  if (-not (Test-Path $FilePath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Token API Tempo manquant.`n`nGenerez-le depuis :`n  Tempo > Settings > API Integration",
      "Credentials Tempo", [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $tempoTokenSecure = Read-Host "Tempo API Token (Bearer)" -AsSecureString
    @{ TempoTokenSecureString=$tempoTokenSecure } | Export-Clixml -Path $FilePath
  }

  $rawContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
  $token      = ""

  if ($rawContent.TrimStart().StartsWith("<")) {
    try {
      $data = Import-Clixml -Path $FilePath
      if ($data.TempoTokenSecureString) {
        $token = [System.Net.NetworkCredential]::new("", $data.TempoTokenSecureString).Password
      } elseif ($data.Token) {
        $token = [string]$data.Token
      }
    } catch {}
  } else {
    $token = $rawContent.Trim()
  }

  if ([string]::IsNullOrWhiteSpace($token)) {
    Log "ERREUR : impossible de lire le token Tempo depuis $FilePath" "ERROR"
    throw "Token Tempo introuvable dans $FilePath"
  }

  return @{
    BaseUrl = "https://api.eu.tempo.io"
    Headers = @{ Authorization="Bearer $token"; Accept="application/json" }
  }
}

# ============================================================
# FILTRE DOMAINES EMAIL
# ============================================================

$allowedDomains = @(
    "mutex.fr", "mutex-exterieur.fr", "harmonie-mutuelle.fr",
    "prestataire.sihm.fr", "chorum.fr"
)

function Get-EmailDomain([string]$email) {
  if ([string]::IsNullOrWhiteSpace($email)) { return $null }
  $idx = $email.IndexOf("@")
  if ($idx -lt 0) { return $null }
  return $email.Substring($idx + 1).Trim().ToLower()
}

function Test-EmailDomainAllowed([string]$email) {
  if ([string]::IsNullOrWhiteSpace($email)) { return $false }
  $domain = Get-EmailDomain $email
  if (-not $domain) { return $false }
  foreach ($d in $script:allowedDomains) {
    if ($domain -eq $d.ToLower()) { return $true }
  }
  return $false
}

# ============================================================
# BANNIERE
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  EXPORT USERS / GROUPES / TEAMS v1.4"  -ForegroundColor Cyan
Write-Host ("  Site  : {0}" -f $SiteOnly)            -ForegroundColor Cyan
Write-Host "  Mode  : Lecture seule"                  -ForegroundColor Green
Write-Host "  Tempo : api.eu.tempo.io"                -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "================================================================"
Log "  EXPORT USERS / GROUPES / TEAMS v1.4"
Log ("  Site : {0}" -f $SiteOnly)
Log "================================================================"

# ============================================================
# ETAPE 1 : CREDENTIALS
# ============================================================

Log "=== ETAPE 1 : Chargement des credentials ==="

$site  = Import-SiteCredentials -FilePath (Join-Path $SecretsDir "site-admin.xml") -SiteName $SiteOnly
$tempo = Import-TempoCredentials -FilePath (Join-Path $SecretsDir "tempo-token")

Log ("  Site Jira : {0}" -f $site.BaseUrl)
Log ("  Tempo API : {0}" -f $tempo.BaseUrl)

# ============================================================
# ETAPE 2 : LISTE DES UTILISATEURS (via groupe Jira)
# ============================================================

Log "=== ETAPE 2 : Chargement des utilisateurs ==="

# Chargement token org pour Fallback 2 (email private)
$orgData    = Import-Clixml -Path (Join-Path $SecretsDir "org-admin.xml")
$orgToken   = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password
$orgHeaders = @{ Authorization="Bearer $orgToken"; Accept="application/json" }

$jiraGroup    = "HM_allusers_jira-software-jiradot"
$groupEncoded = [Uri]::EscapeDataString($jiraGroup)

$activeUsers  = New-Object System.Collections.Generic.List[object]
$seenIds      = @{}
$cSkipInactif = 0
$cSkipDomaine = 0
$cSansEmail   = 0
$startAt      = 0
$pageSize     = 50
$pageNum      = 0
$totalScanned = 0

Log ("  Source : groupe '{0}'" -f $jiraGroup)

while ($true) {
  $pageNum++
  Show-Progress "Etape 2 : Chargement users" `
    ("Page {0} — scannes: {1} — retenus: {2}" -f $pageNum, $totalScanned, $activeUsers.Count) `
    ([Math]::Min($totalScanned, 1100)) 1100

  $url  = "{0}/rest/api/3/group/member?groupname={1}&startAt={2}&maxResults={3}&includeInactiveUsers=false" -f `
            $site.BaseUrl, $groupEncoded, $startAt, $pageSize
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

  if (-not $resp.ok) {
    Log ("  Erreur group/member page {0} : status={1}" -f $pageNum, $resp.status) "ERROR"
    break
  }

  $json   = $resp.content | ConvertFrom-Json
  $values = $json.values
  $count  = if ($values) { ($values | Measure-Object).Count } else { 0 }
  if ($count -eq 0) { break }
  $totalScanned += $count

  foreach ($u in $values) {
    $accId = [string]$u.accountId
    if ($seenIds.ContainsKey($accId)) { continue }
    $seenIds[$accId] = $true

    if (-not [bool]$u.active) { $cSkipInactif++; continue }

    $email = if ($u.emailAddress) { [string]$u.emailAddress } else { "" }

    # --- Fallback 1 : appel direct /user?accountId= ---
    if ([string]::IsNullOrWhiteSpace($email)) {
      $uUrl  = "{0}/rest/api/3/user?accountId={1}" -f $site.BaseUrl, $accId
      $uResp = Invoke-ApiCall -Method "GET" -Url $uUrl -Headers $site.Headers
      if ($uResp.ok) {
        $uData = $uResp.content | ConvertFrom-Json
        if ($uData.emailAddress) { $email = [string]$uData.emailAddress }
      }
    }

    # --- Fallback 2 : GET /manage/profile avec token org (email "private") ---
    if ([string]::IsNullOrWhiteSpace($email)) {
      $adminUrl  = "https://api.atlassian.com/users/{0}/manage/profile" -f $accId
      $adminResp = Invoke-ApiCall -Method "GET" -Url $adminUrl -Headers $orgHeaders
      if ($adminResp.ok) {
        $adminData = $adminResp.content | ConvertFrom-Json
        if ($adminData.account -and $adminData.account.email) {
          $email = [string]$adminData.account.email
          Log ("  ORG-API email recupere : {0} | {1} | {2}" -f $accId, [string]$u.displayName, $email) "DEBUG"
        }
      }
    }

    # --- Log si toujours sans email apres les 3 tentatives ---
    if ([string]::IsNullOrWhiteSpace($email)) {
      Log ("  SANS-EMAIL : {0} | {1}" -f $accId, [string]$u.displayName) "DEBUG"
      $cSansEmail++
    }

    # --- Filtre domaine (uniquement si email non vide) ---
    if (-not [string]::IsNullOrWhiteSpace($email) -and -not (Test-EmailDomainAllowed $email)) {
      Log ("  HORS-DOMAINE : {0} | {1} | {2}" -f $accId, [string]$u.displayName, $email) "DEBUG"
      $cSkipDomaine++
      continue
    }

    $activeUsers.Add(@{
      accountId   = $accId
      displayName = [string]$u.displayName
      email       = $email
    }) | Out-Null
  }

  Log ("  Page {0} : +{1} scannes, {2} retenus au total" -f $pageNum, $count, $activeUsers.Count)

  if ($count -lt $pageSize) { break }
  $startAt += $pageSize
  Start-Sleep -Milliseconds 200
}

Write-Progress -Activity "Etape 2" -Completed
Log ("  Total scannes          : {0}" -f $totalScanned)
Log ("  Retenus                : {0}" -f $activeUsers.Count)
Log ("  Skip inactifs          : {0}" -f $cSkipInactif)
Log ("  Sans email (conserves) : {0}" -f $cSansEmail)
Log ("  Skip hors domaine      : {0}" -f $cSkipDomaine)

# ============================================================
# ETAPE 3 : EQUIPES + MEMBRES TEMPO (pre-chargement)
# ============================================================

Log "=== ETAPE 3 : Pre-chargement equipes + membres Tempo ==="

$tempoTeams       = @{}
$tempoMemberships = New-Object System.Collections.Generic.List[object]
$tempoByUser      = @{}
$tempoApiOk       = $true

# --- 3a : Liste des equipes ---
Log "  3a : Liste des equipes..."
$offset    = 0; $limit = 50; $teamPage = 0
$teamTotal = 0

while ($true) {
  $teamPage++
  Show-Progress "Etape 3a : Equipes Tempo" `
    ("Page {0} — {1} equipes" -f $teamPage, $tempoTeams.Count) `
    ([Math]::Min($tempoTeams.Count, [Math]::Max(1, $teamTotal))) ([Math]::Max(1, $teamTotal))

  $url  = "{0}/4/teams?offset={1}&limit={2}" -f $tempo.BaseUrl, $offset, $limit
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $tempo.Headers -TimeoutSec $TempoTimeoutSec

  if (-not $resp.ok) {
    Log ("  ERREUR : Tempo /4/teams inaccessible (status={0}) — Tempo desactive" -f $resp.status) "ERROR"
    $tempoApiOk = $false; break
  }

  $json         = $resp.content | ConvertFrom-Json
  $teamsOnPage  = $null
  if ($json.results)         { $teamsOnPage = $json.results }
  elseif ($json -is [array]) { $teamsOnPage = $json }

  $fetchedTeams = if ($teamsOnPage) { ($teamsOnPage | Measure-Object).Count } else { 0 }
  if ($fetchedTeams -eq 0) { break }

  if ($json.metadata -and $json.metadata.count) { $teamTotal = [int]$json.metadata.count }

  foreach ($team in $teamsOnPage) {
    $teamId = [string]$team.id
    if (-not $tempoTeams.ContainsKey($teamId)) { $tempoTeams[$teamId] = [string]$team.name }
  }

  Log ("    Page {0} : {1} equipes (+{2})" -f $teamPage, $tempoTeams.Count, $fetchedTeams)

  $isLastTeamPage = $fetchedTeams -lt $limit
  if (-not $isLastTeamPage -and $json.metadata) {
    $isLastTeamPage = -not $json.metadata.next -and -not $json.metadata.hasMore
  }
  if ($isLastTeamPage) { break }

  $offset += $limit
  Start-Sleep -Milliseconds 300
}

Write-Progress -Activity "Etape 3a" -Completed
Log ("  {0} equipes Tempo trouvees" -f $tempoTeams.Count)

# --- 3b : Membres de chaque equipe ---
if ($tempoApiOk -and $tempoTeams.Count -gt 0) {
  Log "  3b : Chargement des membres par equipe..."

  $teamList  = @($tempoTeams.Keys)
  $teamTotal = $teamList.Count
  $teamIdx   = 0

  foreach ($teamId in $teamList) {
    $teamIdx++
    $teamName = $tempoTeams[$teamId]

    Show-Progress "Etape 3b : Membres Tempo" `
      ("{0}/{1} — {2} (total membres: {3})" -f $teamIdx, $teamTotal, $teamName, $tempoMemberships.Count) `
      $teamIdx $teamTotal

    Log ("    [{0}/{1}] {2}" -f $teamIdx, $teamTotal, $teamName)

    $mOffset     = 0
    $memberPage  = 0
    $memberTotal = 0
    $maxPages    = 50

    while ($true) {
      $memberPage++
      if ($memberPage -gt $maxPages) {
        Log ("    WARN : {0} depasse {1} pages — arret" -f $teamName, $maxPages) "WARN"
        break
      }

      $mUrl  = "{0}/4/team-memberships/team/{1}?offset={2}&limit=50" -f $tempo.BaseUrl, $teamId, $mOffset
      $mResp = Invoke-ApiCall -Method "GET" -Url $mUrl -Headers $tempo.Headers -TimeoutSec $TempoTimeoutSec

      if (-not $mResp.ok) {
        Log ("    WARN : membres {0} inaccessibles (status={1})" -f $teamName, $mResp.status) "WARN"
        break
      }

      $mJson   = $mResp.content | ConvertFrom-Json
      $members = $null
      if ($mJson.results)         { $members = $mJson.results }
      elseif ($mJson -is [array]) { $members = $mJson }

      $fetchedM = if ($members) { ($members | Measure-Object).Count } else { 0 }
      if ($fetchedM -eq 0) { break }

      foreach ($m in $members) {
        $accId = ""
        if ($m.member -and $m.member.accountId) { $accId = [string]$m.member.accountId }
        elseif ($m.accountId)                   { $accId = [string]$m.accountId }
        if ([string]::IsNullOrWhiteSpace($accId)) { continue }

        $dateFrom = ""; $dateTo = ""; $role = ""
        if ($m.from)         { $dateFrom = [string]$m.from }
        elseif ($m.dateFrom) { $dateFrom = [string]$m.dateFrom }
        if ($m.to)           { $dateTo   = [string]$m.to }
        elseif ($m.dateTo)   { $dateTo   = [string]$m.dateTo }
        if ($m.role -and $m.role.name) { $role = [string]$m.role.name }
        elseif ($m.roleName)           { $role = [string]$m.roleName }

        $tempoMemberships.Add(@{
          teamId    = $teamId
          teamName  = $teamName
          accountId = $accId
          dateFrom  = $dateFrom
          dateTo    = $dateTo
          role      = $role
        }) | Out-Null
        $memberTotal++
      }

      if ($memberTotal -gt 50) {
        Log ("      page {0} : {1} membres pour {2}" -f $memberPage, $memberTotal, $teamName)
      }

      $isLastM = $fetchedM -lt 50
      if (-not $isLastM -and $mJson.metadata) {
        $isLastM = -not $mJson.metadata.next -and -not $mJson.metadata.hasMore
      }
      if (-not $isLastM -and $mJson.total) {
        $isLastM = $memberTotal -ge [int]$mJson.total
      }
      if ($isLastM) { break }

      $mOffset += 50
      Start-Sleep -Milliseconds 150
    }

    Log ("    -> {0} membres pour {1}" -f $memberTotal, $teamName)
  }

  Write-Progress -Activity "Etape 3b" -Completed
  Log ("  {0} memberships Tempo au total" -f $tempoMemberships.Count)

} else {
  Log "  Etape 3b ignoree (Tempo indisponible ou 0 equipes)" "WARN"
}

# --- Index par accountId ---
$tempoByUser = @{}
foreach ($m in $tempoMemberships) {
  if (-not $tempoByUser.ContainsKey($m.accountId)) {
    $tempoByUser[$m.accountId] = New-Object System.Collections.Generic.List[object]
  }
  $tempoByUser[$m.accountId].Add($m) | Out-Null
}
Log ("  {0} users avec au moins une equipe Tempo" -f $tempoByUser.Count)
# ============================================================
# ETAPE 4 : BOUCLE PRINCIPALE
# ============================================================

Log "=== ETAPE 4 : Groupes + equipes + licence Tempo par utilisateur ==="

$exportResults  = New-Object System.Collections.Generic.List[object]
$cProcessed     = 0
$cTempoLicensed = 0
$cCount         = $activeUsers.Count
$startTime      = Get-Date
$logEvery       = [Math]::Max(1, [int]($cCount / 20))

foreach ($user in $activeUsers) {
  $cProcessed++

  $elapsed   = ((Get-Date) - $startTime).TotalSeconds
  $rate      = if ($elapsed -gt 0 -and $cProcessed -gt 1) { $cProcessed / $elapsed } else { 1 }
  $remaining = if ($rate -gt 0) { [int](($cCount - $cProcessed) / $rate) } else { 0 }
  $eta       = "{0:mm\:ss}" -f [TimeSpan]::FromSeconds($remaining)

  Show-Progress "Etape 4 : Traitement users" `
    ("{0}/{1} — {2} — ETA {3} — Tempo: {4}" -f `
      $cProcessed, $cCount, $user.displayName, $eta, $cTempoLicensed) `
    $cProcessed $cCount

  if ($cProcessed % $logEvery -eq 0 -or $cProcessed -eq 1 -or $cProcessed -eq $cCount) {
    Log ("  [{0}/{1}] {2} — ETA {3}" -f $cProcessed, $cCount, $user.displayName, $eta)
  }

  $accId = $user.accountId

  # --- Groupes Jira ---
  $groupNames = @()
  $url  = "{0}/rest/api/3/user/groups?accountId={1}" -f $site.BaseUrl, $accId
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers
  if ($resp.ok) {
    $groupNames = @(($resp.content | ConvertFrom-Json) |
      ForEach-Object { [string]$_.name } | Sort-Object)
  }
  $groupesStr = if ($groupNames.Count -gt 0) { $groupNames -join " | " } else { "" }

  # --- Equipes Tempo ---
  $tempoTeamsStr = ""; $dernEquipe = ""; $dernDateFrom = ""; $dernDateTo = ""; $dernRole = ""
  $nbTempoTeams  = 0

  if ($tempoByUser -and $tempoByUser.ContainsKey($accId)) {
    $memberships   = $tempoByUser[$accId]
    $teamNames     = @($memberships | ForEach-Object { $_.teamName } |
                       Where-Object { $_ } | Sort-Object -Unique)
    $nbTempoTeams  = $teamNames.Count
    $tempoTeamsStr = $teamNames -join " | "

    $sorted = $memberships | Sort-Object {
      if ($_.dateFrom) { try { [DateTime]::Parse($_.dateFrom) } catch { [DateTime]::MinValue } }
      else { [DateTime]::MinValue }
    } -Descending

    $lastActive = $sorted |
      Where-Object { [string]::IsNullOrWhiteSpace($_.dateTo) } |
      Select-Object -First 1
    if (-not $lastActive) { $lastActive = $sorted | Select-Object -First 1 }

    if ($lastActive) {
      $dernEquipe   = $lastActive.teamName
      $dernDateFrom = Format-DateFr $lastActive.dateFrom
      $dernDateTo   = if ($lastActive.dateTo) { Format-DateFr $lastActive.dateTo } else { "(en cours)" }
      $dernRole     = $lastActive.role
    }
  }

  # --- Licence Tempo (via user-schedule) ---
  $tempoLicensed = $false
  $tempoRoleStr  = if ($dernRole) { $dernRole } else { "" }

  if ($tempoApiOk) {
    $schedUrl  = "{0}/4/user-schedule/{1}?from={2}&to={3}" -f $tempo.BaseUrl, $accId,
                   (Get-Date).ToString("yyyy-MM-dd"), (Get-Date).AddDays(1).ToString("yyyy-MM-dd")
    $schedResp = Invoke-ApiCall -Method "GET" -Url $schedUrl -Headers $tempo.Headers -TimeoutSec $TempoTimeoutSec
    if ($schedResp.ok) {
      $tempoLicensed = $true
      $cTempoLicensed++
    }
  }

  # --- Compilation ---
  $exportResults.Add(@{
    accountId       = $accId
    displayName     = $user.displayName
    email           = $user.email
    groupes         = $groupesStr
    nbGroupes       = [int]$groupNames.Count
    tempoTeams      = $tempoTeamsStr
    nbTempoTeams    = [int]$nbTempoTeams
    dernEquipeTempo = $dernEquipe
    dernDateFrom    = $dernDateFrom
    dernDateTo      = $dernDateTo
    dernRole        = $tempoRoleStr
    tempoLicence    = $tempoLicensed
  }) | Out-Null

  Start-Sleep -Milliseconds 50
}

Write-Progress -Activity "Etape 4" -Completed
Log ("  {0} utilisateurs traites, {1} avec licence Tempo" -f $cProcessed, $cTempoLicensed)

# ============================================================
# ETAPE 5 : EXPORT CSV
# ============================================================

Log "=== ETAPE 5 : Export CSV ==="

$csvColumns = @(
  "AccountId","DisplayName","Email",
  "NbGroupes","Groupes",
  "NbEquipesTempo","EquipesTempo",
  "DernEquipeTempo","DernDateFrom","DernDateTo","DernRole",
  "LicenceTempo"
)

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$sb      = New-Object System.Text.StringBuilder
[void]$sb.AppendLine(($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";")

foreach ($r in ($exportResults | Sort-Object { $_.displayName })) {
  $row = [ordered]@{
    AccountId       = $r.accountId
    DisplayName     = $r.displayName
    Email           = $r.email
    NbGroupes       = $r.nbGroupes
    Groupes         = $r.groupes
    NbEquipesTempo  = $r.nbTempoTeams
    EquipesTempo    = $r.tempoTeams
    DernEquipeTempo = $r.dernEquipeTempo
    DernDateFrom    = $r.dernDateFrom
    DernDateTo      = $r.dernDateTo
    DernRole        = $r.dernRole
    LicenceTempo    = if ($r.tempoLicence) { "OUI" } else { "NON" }
  }
  [void]$sb.AppendLine(($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";")
}

[System.IO.File]::WriteAllText($csvFile, $sb.ToString(), $utf8Bom)
Log ("  CSV -> {0}" -f $csvFile)

# ============================================================
# ETAPE 6 : AFFICHAGE CONSOLE
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  {0} UTILISATEURS TRAITES" -f $exportResults.Count) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$withTempo   = ($exportResults | Where-Object { $_.nbTempoTeams -gt 0 } | Measure-Object).Count
$withLicence = ($exportResults | Where-Object { $_.tempoLicence } | Measure-Object).Count
$withNoEmail = ($exportResults | Where-Object { [string]::IsNullOrWhiteSpace($_.email) } | Measure-Object).Count

$totalGroupes = 0
foreach ($r in $exportResults) { $totalGroupes += [int]$r.nbGroupes }
$avgGroups = if ($exportResults.Count -gt 0) {
  [Math]::Round($totalGroupes / $exportResults.Count, 1)
} else { 0 }

Write-Host ("  Avec equipe Tempo       : {0}" -f $withTempo)        -ForegroundColor Yellow
Write-Host ("  Avec licence Tempo      : {0}" -f $withLicence)      -ForegroundColor Yellow
Write-Host ("  Sans email (conserves)  : {0}" -f $withNoEmail)      -ForegroundColor DarkYellow
Write-Host ("  Moyenne groupes / user  : {0}" -f $avgGroups)        -ForegroundColor DarkGray
Write-Host ("  Equipes Tempo totales   : {0}" -f $tempoTeams.Count) -ForegroundColor Cyan
Write-Host ("  Memberships Tempo       : {0}" -f $tempoMemberships.Count) -ForegroundColor DarkGray
Write-Host ""

$displayMax  = [Math]::Min(30, $exportResults.Count)
$idx         = 0
$displayList = $exportResults | Sort-Object { $_.displayName }

foreach ($r in $displayList) {
  $idx++
  if ($idx -gt $displayMax) {
    Write-Host ("  ... et {0} autres (voir CSV)" -f ($exportResults.Count - $displayMax)) -ForegroundColor DarkGray
    break
  }

  $nameColor = if ($r.tempoLicence) { "Cyan" } else { "White" }
  $groupDisp = if ($r.groupes) {
    if ($r.groupes.Length -gt 80) { $r.groupes.Substring(0,80)+"..." } else { $r.groupes }
  } else { "-" }

  Write-Host ("{0,4}. {1}" -f $idx, $r.displayName) -ForegroundColor $nameColor
  if ([string]::IsNullOrWhiteSpace($r.email)) {
    Write-Host ("      Email         : (vide)" ) -ForegroundColor DarkYellow
  }
  Write-Host ("      Groupes ({0})   : {1}" -f $r.nbGroupes, $groupDisp) -ForegroundColor DarkGray
  if ($r.tempoTeams) {
    Write-Host ("      Tempo ({0})    : {1}" -f $r.nbTempoTeams, $r.tempoTeams) -ForegroundColor DarkCyan
    Write-Host ("      Derniere      : {0}  ({1} -> {2})  [{3}]" -f `
      $r.dernEquipeTempo, $r.dernDateFrom, $r.dernDateTo, $r.dernRole) -ForegroundColor DarkCyan
  }
  $licenceColor = if ($r.tempoLicence) { "Green" } else { "DarkGray" }
  Write-Host ("      Licence Tempo : {0}" -f $(if ($r.tempoLicence) { "OUI" } else { "NON" })) -ForegroundColor $licenceColor
  Write-Host ""
}

# ============================================================
# RESUME FINAL
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Utilisateurs traites    : {0}" -f $exportResults.Count)
Write-Host ("  Avec equipe Tempo       : {0}" -f $withTempo)              -ForegroundColor Yellow
Write-Host ("  Avec licence Tempo      : {0}" -f $withLicence)            -ForegroundColor Yellow
Write-Host ("  Sans email (conserves)  : {0}" -f $withNoEmail)            -ForegroundColor DarkYellow
Write-Host ("  Moyenne groupes / user  : {0}" -f $avgGroups)              -ForegroundColor DarkGray
Write-Host ("  Equipes Tempo totales   : {0}" -f $tempoTeams.Count)       -ForegroundColor Cyan
Write-Host ("  Memberships Tempo       : {0}" -f $tempoMemberships.Count) -ForegroundColor DarkGray
Write-Host ("  Duree                   : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  Horodatage              : {0}" -f (Get-Date).ToString("dd/MM/yyyy HH:mm:ss"))
Write-Host ""
Write-Host "  Fichiers :" -ForegroundColor White
Write-Host ("    CSV : {0}" -f $csvFile)
Write-Host ("    LOG : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."