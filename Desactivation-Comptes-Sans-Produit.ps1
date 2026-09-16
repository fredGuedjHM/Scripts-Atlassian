<#
.SYNOPSIS
  Desactivation-Comptes-Sans-Produit.ps1
  Identifie et desactive les comptes Atlassian sans acces produit
  sur les domaines Jiradot et mutexfr, crees/invites depuis plus de 2 mois.

.DESCRIPTION
  Ce script :
  1. Pre-charge les utilisateurs des sites Jiradot et mutexfr
  2. Detecte DYNAMIQUEMENT les groupes Jira donnant acces produit
     via l API /rest/api/3/applicationrole
  3. Parcourt les utilisateurs manages de l organisation via l Admin API
     et verifie l acces produit via l API profil
  4. Recupere les groupes d habilitation des pre-candidats
  5. Exclut les comptes ayant un acces produit Jira via groupe (applicationrole)
  6. Verifie l acces CONFLUENCE directement via l API Confluence
     (car Confluence n apparait PAS dans applicationrole)
  7. Exclut les comptes qui ne sont QUE customers JSM
  8. Mode Dry    : affiche la liste des candidats avec groupes (aucune action)
  9. Mode Execute : desactive au cas par cas avec confirmation individuelle

  TRIPLE VERIFICATION ACCES PRODUIT :
    1. API /users/{id}/manage/profile → product_access
    2. Appartenance a un groupe Jira produit (via applicationrole)
    3. Acces Confluence reel (via API wiki/rest/api/user)

  POURQUOI TRIPLE VERIFICATION :
    - L API product_access ne reflete pas toujours l acces via groupes
    - L API applicationrole ne couvre QUE Jira (pas Confluence, Loom, etc.)
    - L API Confluence repond directement si le user a acces ou non

  CREDENTIALS :
    secrets\org-admin.xml    => OrgId + API Key
    secrets\site-admin.xml   => Jiradot
    secrets\site-mutexfr.xml => mutexfr

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 4.5 — Juillet 2026
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Desactivation-Comptes-Sans-Produit.ps1
  .\Desactivation-Comptes-Sans-Produit.ps1 -Mode Execute
  .\Desactivation-Comptes-Sans-Produit.ps1 -SeuilMois 3
#>

[CmdletBinding()]
param(
  [ValidateSet("Dry","Execute")]
  [string] $Mode = "Dry",

  [int] $SeuilMois = 2,

  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("ComptesSansProduit_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("ComptesSansProduit_{0}.csv" -f $ts)
$csvRolesFile = Join-Path $ExportsDir ("GroupesProduit-ApplicationRoles_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Write-CsvLine([string]$line) {
  [System.IO.File]::AppendAllText($script:csvFile, "$line`r`n", [System.Text.Encoding]::UTF8)
}

function Write-CsvHeader([string]$line) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($script:csvFile, "$line`r`n", $utf8Bom)
}

function Format-DateFr([string]$isoDate) {
  if (-not $isoDate) { return "" }
  try {
    $dt = [DateTimeOffset]::Parse($isoDate)
    return $dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
  } catch { return $isoDate }
}

# ============================================================
# FILTRE DOMAINES EMAIL
# ============================================================

$allowedDomains = @(
    "mutex.fr",
    "mutex-exterieur.fr",
    "harmonie-mutuelle.fr",
    "prestataire.sihm.fr",
    "chorum.fr"
)

function Get-EmailDomain([string]$email) {
    if ([string]::IsNullOrWhiteSpace($email)) { return $null }
    $idx = $email.IndexOf("@")
    if ($idx -lt 0) { return $null }
    return $email.Substring($idx + 1).Trim().ToLower()
}

function Test-EmailDomainAllowed([string]$email) {
    if ([string]::IsNullOrWhiteSpace($email)) { return $true }
    $domain = Get-EmailDomain $email
    if (-not $domain) { return $true }
    foreach ($d in $script:allowedDomains) {
        if ($domain -eq $d.ToLower()) { return $true }
    }
    return $false
}

# ============================================================
# GROUPES CUSTOMER (utilisateurs legitimes sans licence)
# ============================================================

$customerGroupPatterns = @(
    "servicemanagement-customers",
    "jira-servicedesk-users",
    "service-desk-customers",
    "jsm-customers"
)

function Test-IsCustomerOnly([string]$groupesStr) {
    if ([string]::IsNullOrWhiteSpace($groupesStr)) { return $false }
    $groupList = $groupesStr -split '\s*\|\s*'
    foreach ($g in $groupList) {
        $gTrimmed = $g.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($gTrimmed)) { continue }
        $isCustomerGroup = $false
        foreach ($pattern in $customerGroupPatterns) {
            if ($gTrimmed.Contains($pattern.ToLower())) {
                $isCustomerGroup = $true
                break
            }
        }
        if (-not $isCustomerGroup) { return $false }
    }
    return $true
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
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{ Method=$Method; Uri=$Url; Headers=$Headers; UseBasicParsing=$true; ErrorAction="Stop" }
      if ($Body) { $params["ContentType"]="application/json; charset=utf-8"; $params["Body"]=[System.Text.Encoding]::UTF8.GetBytes($Body) }
      $resp = Invoke-WebRequest @params
      $contentUtf8 = $resp.Content
      try {
        $stream = $resp.RawContentStream
        if ($stream -and $stream.CanSeek) {
          $stream.Position = 0
          $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
          $contentUtf8 = $reader.ReadToEnd(); $reader.Close()
        }
      } catch {
        try {
          $isoBytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($resp.Content)
          $contentUtf8 = [System.Text.Encoding]::UTF8.GetString($isoBytes)
        } catch {}
      }
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$contentUtf8 }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $rd = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $rd.ReadToEnd(); $rd.Close()
      } catch {}
      if ($attempt -gt $MaxRetries) { return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody } }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(5, $attempt)))
        Log ("Retry {0} status={1} in {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
        Start-Sleep -Seconds $sleepSec; continue
      }
      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# ============================================================
# CREDENTIALS
# ============================================================

function Load-SiteCredentials([string]$CredFile, [string]$SiteName) {
  if (-not (Test-Path $CredFile)) {
    Log "Fichier $CredFile introuvable, creation interactive..." "WARN"
    [System.Windows.Forms.MessageBox]::Show(
      ("Credentials pour {0} manquants.`n`nVous allez fournir :`n  1. URL du site`n  2. Email admin`n  3. API Token" -f $SiteName),
      "Credentials $SiteName", [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl = Read-Host "URL du site (ex: $SiteName.atlassian.net)"
    $inputUrl = $inputUrl -replace "^https?://", "" -replace "/.*$", "" -replace "/$", ""
    $adminEmail = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $CredFile
    Log "Credentials sauvegardes dans $CredFile"
  }
  $data = Import-Clixml -Path $CredFile
  $url = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
    Name    = $SiteName
  }
}

function Get-SiteUsers([hashtable]$Site) {
  $usersMap = @{}
  $startAt = 0; $pageSize = 200
  while ($true) {
    $url = "{0}/rest/api/3/users/search?startAt={1}&maxResults={2}" -f $Site.BaseUrl, $startAt, $pageSize
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    if (-not $resp.ok) {
      Log ("  Erreur listing users {0} : status={1}" -f $Site.Name, $resp.status) "ERROR"
      break
    }
    $json = $resp.content | ConvertFrom-Json
    $count = ($json | Measure-Object).Count
    if ($count -eq 0) { break }
    foreach ($u in $json) {
      $accId = [string]$u.accountId
      $accType = ""; if ($u.accountType) { $accType = [string]$u.accountType }
      if ($accType -eq "app") { continue }
      if (-not $usersMap.ContainsKey($accId)) {
        $usersMap[$accId] = @{
          displayName = [string]$u.displayName
          active      = [bool]$u.active
          accountType = $accType
        }
      }
    }
    if ($count -lt $pageSize) { break }
    $startAt += $pageSize
    Start-Sleep -Milliseconds 200
  }
  Log ("  {0} : {1} utilisateurs charges (hors apps)" -f $Site.Name, $usersMap.Count)
  return $usersMap
}

# ============================================================
# DETECTION DYNAMIQUE DES GROUPES PRODUIT (applicationrole)
# ============================================================

function Get-ApplicationRoles([hashtable]$Site) {
    $roles = New-Object System.Collections.Generic.List[object]
    $url = "{0}/rest/api/3/applicationrole" -f $Site.BaseUrl
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers

    if (-not $resp.ok) {
        Log ("  WARN : applicationrole indisponible sur {0} (status={1})" -f $Site.Name, $resp.status) "WARN"
        return $roles
    }

    $json = $resp.content | ConvertFrom-Json

    foreach ($role in $json) {
        $roleKey  = ""; if ($role.key)  { $roleKey  = [string]$role.key }
        $roleName = ""; if ($role.name) { $roleName = [string]$role.name }

        if ($role.groups) {
            foreach ($g in $role.groups) {
                $gName = [string]$g
                if ([string]::IsNullOrWhiteSpace($gName)) { continue }
                $isDefault = $false
                if ($role.defaultGroups) {
                    $isDefault = ($role.defaultGroups | ForEach-Object { [string]$_ }) -contains $gName
                }
                $roles.Add(@{
                    RoleKey   = $roleKey
                    RoleName  = $roleName
                    GroupName = $gName
                    IsDefault = $isDefault
                    Site      = $Site.Name
                }) | Out-Null
            }
        }

        if ($role.defaultGroups) {
            foreach ($g in $role.defaultGroups) {
                $gName = [string]$g
                if ([string]::IsNullOrWhiteSpace($gName)) { continue }
                $alreadyAdded = $roles | Where-Object { $_.GroupName -eq $gName -and $_.Site -eq $Site.Name -and $_.RoleKey -eq $roleKey }
                if (-not $alreadyAdded) {
                    $roles.Add(@{
                        RoleKey   = $roleKey
                        RoleName  = $roleName
                        GroupName = $gName
                        IsDefault = $true
                        Site      = $Site.Name
                    }) | Out-Null
                }
            }
        }
    }

    return $roles
}

# ============================================================
# VERIFICATION ACCES CONFLUENCE (API directe)
# ============================================================

function Test-ConfluenceAccess([hashtable]$Site, [string]$AccountId) {
    <#
    Verifie si un utilisateur a acces a Confluence sur ce site.
    Appelle GET /wiki/rest/api/user?accountId={id}
      - 200 → acces Confluence confirme
      - 401/403/404 → pas d acces
      - Autre erreur → on ne sait pas (retourne $false par prudence)
    #>
    $url = "{0}/wiki/rest/api/user?accountId={1}" -f $Site.BaseUrl, $AccountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers

    if ($resp.ok) {
        return $true
    }

    # 401/403/404 = pas d acces → normal
    if ($resp.status -eq 401 -or $resp.status -eq 403 -or $resp.status -eq 404) {
        return $false
    }

    # Autre erreur (500, timeout, etc.) → on ne bloque pas, log + retourne false
    Log ("    Confluence check erreur pour {0} sur {1} : status={2}" -f $AccountId, $Site.Name, $resp.status) "DEBUG"
    return $false
}

# ============================================================
# CONFIGURATION
# ============================================================

$dateSeuil = (Get-Date).AddMonths(-$SeuilMois)
$dateSeuilStr = $dateSeuil.ToString("dd/MM/yyyy")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  DESACTIVATION COMPTES SANS PRODUIT" -ForegroundColor Cyan
Write-Host ("  Mode: {0} | Seuil: {1} mois (avant le {2})" -f $Mode, $SeuilMois, $dateSeuilStr) -ForegroundColor $(if ($Mode -eq "Dry") { "Green" } else { "Red" })
Write-Host "  Domaines : Jiradot + mutexfr" -ForegroundColor Cyan
Write-Host "  Emails   : $($allowedDomains -join ', ')" -ForegroundColor Cyan
Write-Host "  Detection produit : API + applicationrole + Confluence directe" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "================================================================"
Log "  DESACTIVATION COMPTES SANS PRODUIT"
Log ("  Mode: {0} | Seuil: {1} mois (avant le {2})" -f $Mode, $SeuilMois, $dateSeuilStr)
Log "  Domaines cibles : Jiradot + mutexfr"
Log ("  Emails autorises : {0}" -f ($allowedDomains -join ", "))
Log "  Detection : API product_access + applicationrole (Jira) + API Confluence directe"
Log ("  Patterns groupes customer : {0}" -f ($customerGroupPatterns -join ", "))
Log "================================================================"

# ============================================================
# ETAPE 1 : CREDENTIALS
# ============================================================

Log "=== ETAPE 1 : Chargement des credentials ==="

$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) {
  Log "Fichier org-admin.xml introuvable, creation interactive..." "WARN"
  [System.Windows.Forms.MessageBox]::Show(
    ("Le fichier org-admin.xml n'existe pas.`n`nVous allez fournir :`n  1. L'OrgId`n  2. Une API Key d'organisation"),
    "Credentials Organisation", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  $inputOrgId = Read-Host "OrgId (visible dans admin.atlassian.com)"
  $apiKeySecure = Read-Host "API Key d'organisation" -AsSecureString
  @{ OrgId=$inputOrgId; ApiKeySecureString=$apiKeySecure } | Export-Clixml -Path $orgCredFile
  Log "Credentials sauvegardes dans $orgCredFile"
}
$orgData = Import-Clixml -Path $orgCredFile
$orgId = [string]$orgData.OrgId
$orgApiKey = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password
$orgHeaders = @{ Authorization = "Bearer $orgApiKey"; Accept = "application/json" }

$siteJiradot = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
$siteMutexfr = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-mutexfr.xml") -SiteName "mutexfr"

Log ("  OrgId    : {0}" -f $orgId)
Log ("  Jiradot  : {0}" -f $siteJiradot.BaseUrl)
Log ("  mutexfr  : {0}" -f $siteMutexfr.BaseUrl)

# ============================================================
# ETAPE 2 : PRE-CHARGEMENT DES UTILISATEURS PAR SITE
# ============================================================

Log "=== ETAPE 2 : Pre-chargement des utilisateurs par site ==="

$jiradotUsers = Get-SiteUsers -Site $siteJiradot
$mutexfrUsers = Get-SiteUsers -Site $siteMutexfr

# ============================================================
# ETAPE 2b : DETECTION DYNAMIQUE DES GROUPES PRODUIT JIRA
# ============================================================

Log "=== ETAPE 2b : Detection des groupes Jira donnant acces produit (applicationrole) ==="

$rolesJiradot = Get-ApplicationRoles -Site $siteJiradot
$rolesMutexfr = Get-ApplicationRoles -Site $siteMutexfr

# Fusion
$allRoles = New-Object System.Collections.Generic.List[object]
foreach ($r in $rolesJiradot) { $allRoles.Add($r) | Out-Null }
foreach ($r in $rolesMutexfr) { $allRoles.Add($r) | Out-Null }

# HashSet pour verification rapide
$allProductGroupNames = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
foreach ($r in $allRoles) {
    [void]$allProductGroupNames.Add($r.GroupName)
}

Log ("  Jiradot : {0} associations groupe-role" -f $rolesJiradot.Count)
foreach ($r in ($rolesJiradot | Sort-Object { $_.RoleName + $_.GroupName })) {
    $def = if ($r.IsDefault) { " (default)" } else { "" }
    Log ("    [{0}] {1}{2}" -f $r.RoleName, $r.GroupName, $def) "DEBUG"
}

Log ("  mutexfr : {0} associations groupe-role" -f $rolesMutexfr.Count)
foreach ($r in ($rolesMutexfr | Sort-Object { $_.RoleName + $_.GroupName })) {
    $def = if ($r.IsDefault) { " (default)" } else { "" }
    Log ("    [{0}] {1}{2}" -f $r.RoleName, $r.GroupName, $def) "DEBUG"
}

Log ("  Total groupes Jira produit uniques : {0}" -f $allProductGroupNames.Count)

# --- Export CSV des groupes/roles ---
Log "  Export CSV groupes-roles -> $csvRolesFile"

$rolesCsvColumns = @("Site","ApplicationRole","RoleKey","GroupName","IsDefault")
$rolesCsvHeader = ($rolesCsvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$rolesSb = New-Object System.Text.StringBuilder
[void]$rolesSb.AppendLine($rolesCsvHeader)

foreach ($r in ($allRoles | Sort-Object { $_.Site + $_.RoleName + $_.GroupName })) {
    $row = [ordered]@{
        Site            = $r.Site
        ApplicationRole = $r.RoleName
        RoleKey         = $r.RoleKey
        GroupName       = $r.GroupName
        IsDefault       = if ($r.IsDefault) { "OUI" } else { "NON" }
    }
    $line = ($rolesCsvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
    [void]$rolesSb.AppendLine($line)
}

[System.IO.File]::WriteAllText($csvRolesFile, $rolesSb.ToString(), $utf8Bom)
Log ("  CSV groupes-roles : {0} lignes" -f $allRoles.Count)

# Fonction de verification Jira (comparaison EXACTE via HashSet)
function Test-HasJiraProductGroup([string]$groupesStr) {
    if ([string]::IsNullOrWhiteSpace($groupesStr)) { return $false }
    $groupList = $groupesStr -split '\s*\|\s*'
    foreach ($g in $groupList) {
        $gTrimmed = $g.Trim()
        if ([string]::IsNullOrWhiteSpace($gTrimmed)) { continue }
        if ($allProductGroupNames.Contains($gTrimmed)) { return $true }
    }
    return $false
}

# ============================================================
# ETAPE 3 : BOUCLE PRINCIPALE — ORG API + PROFIL (fusionnees)
# ============================================================

Log "=== ETAPE 3 : Parcours des utilisateurs manages + verification produit (API) ==="

$candidatsBruts = New-Object System.Collections.Generic.List[object]
$cursor = $null
$startTime = Get-Date

# Compteurs
$cTotal           = 0
$cSkipInactive    = 0
$cSkipApp         = 0
$cSkipHorsDomaine = 0
$cSkipHorsEmail   = 0
$cSkipSiteInactif = 0
$cAvecProduitAPI  = 0
$cTropRecent      = 0
$cSansProduitAPI  = 0
$cErreurs         = 0

while ($true) {
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/users?maxResults=100"
  if ($cursor) { $url += "&cursor=$cursor" }

  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $orgHeaders

  if (-not $resp.ok) {
    Log ("Erreur API users : status={0} {1}" -f $resp.status, $resp.error) "ERROR"
    break
  }

  $json = $resp.content | ConvertFrom-Json

  foreach ($user in $json.data) {
    $cTotal++

    $accId       = [string]$user.account_id
    $accType     = ""; if ($user.account_type) { $accType = [string]$user.account_type }
    $accStatus   = ""; if ($user.account_status) { $accStatus = [string]$user.account_status }
    $displayName = ""; if ($user.name) { $displayName = [string]$user.name }
    $email       = ""; if ($user.email) { $email = [string]$user.email }
    $lastActive  = ""; if ($user.last_active) { $lastActive = [string]$user.last_active }

    # --- FILTRE 1 : Statut ---
    if ($accStatus -eq "inactive" -or $accStatus -eq "closed" -or $accStatus -eq "suspended") {
      $cSkipInactive++
      continue
    }

    # --- FILTRE 2 : Type ---
    if ($accType -eq "app") {
      $cSkipApp++
      continue
    }

    # --- FILTRE 3 : Presence sur Jiradot/mutexfr ---
    $onJiradot = $jiradotUsers.ContainsKey($accId)
    $onMutexfr = $mutexfrUsers.ContainsKey($accId)

    if (-not $onJiradot -and -not $onMutexfr) {
      $cSkipHorsDomaine++
      continue
    }

    # --- FILTRE 4 : Email domaine ---
    if ($email -and -not (Test-EmailDomainAllowed $email)) {
      $cSkipHorsEmail++
      continue
    }

    # --- FILTRE 5 : Deja inactif cote site ---
    $alreadyInactive = $false
    if ($onJiradot -and $jiradotUsers[$accId].active -eq $false -and (-not $onMutexfr -or $mutexfrUsers[$accId].active -eq $false)) {
      $alreadyInactive = $true
    }
    if (-not $onJiradot -and $onMutexfr -and $mutexfrUsers[$accId].active -eq $false) {
      $alreadyInactive = $true
    }
    if ($alreadyInactive) {
      $cSkipSiteInactif++
      continue
    }

    # --- FETCH PROFIL ---
    $profileUrl = "https://api.atlassian.com/users/$accId/manage/profile"
    $profResp = Invoke-ApiCall -Method "GET" -Url $profileUrl -Headers $orgHeaders

    if (-not $profResp.ok) {
      if ($profResp.status -eq 404) { continue }
      Log ("  Erreur profil {0} ({1}) : status={2}" -f $displayName, $accId, $profResp.status) "WARN"
      $cErreurs++
      Start-Sleep -Milliseconds 100
      continue
    }

    $profile = $profResp.content | ConvertFrom-Json

    # Extraire email si manquant
    if (-not $email) {
      if ($profile.account -and $profile.account.email) { $email = [string]$profile.account.email }
      elseif ($profile.email) { $email = [string]$profile.email }

      if ($email -and -not (Test-EmailDomainAllowed $email)) {
        $cSkipHorsEmail++
        continue
      }
    }

    # Extraire les dates
    $dateCreation = ""
    $dateInvitation = ""

    if ($profile.account -and $profile.account.created) { $dateCreation = [string]$profile.account.created }
    elseif ($profile.created) { $dateCreation = [string]$profile.created }

    if ($profile.account -and $profile.account.invited) { $dateInvitation = [string]$profile.account.invited }
    elseif ($profile.invited) { $dateInvitation = [string]$profile.invited }

    $dateReference = $dateCreation
    if ($dateInvitation) { $dateReference = $dateInvitation }

    # --- FILTRE 6 : Acces produit (API) ---
    $hasProduct = $false
    $productList = ""

    if ($profile.product_access) {
      foreach ($pa in $profile.product_access) {
        if ($pa.products -and ($pa.products | Measure-Object).Count -gt 0) {
          $hasProduct = $true
          $siteName = ""
          if ($pa.url) { $siteName = [string]$pa.url }
          elseif ($pa.name) { $siteName = [string]$pa.name }
          elseif ($pa.site_name) { $siteName = [string]$pa.site_name }

          $prods = ($pa.products | ForEach-Object {
            $pName = ""
            if ($_.name) { $pName = [string]$_.name }
            elseif ($_.product_name) { $pName = [string]$_.product_name }
            $pName
          }) -join ", "

          if ($productList) { $productList += " | " }
          $productList += ("{0}: {1}" -f $siteName, $prods)
        }
      }
    }

    if ($hasProduct) {
      $cAvecProduitAPI++
      continue
    }

    # --- FILTRE 7 : Date trop recente ---
    $isOldEnough = $false
    if ($dateReference) {
      try {
        $dtRef = [DateTimeOffset]::Parse($dateReference)
        if ($dtRef.DateTime -lt $dateSeuil) { $isOldEnough = $true }
      } catch {}
    } else {
      $isOldEnough = $true
    }

    if (-not $isOldEnough) {
      $cTropRecent++
      continue
    }

    # === PRE-CANDIDAT ===
    $cSansProduitAPI++

    $domaines = @()
    if ($onJiradot) { $domaines += "Jiradot" }
    if ($onMutexfr) { $domaines += "mutexfr" }
    $domainesStr = $domaines -join " + "

    $candidatsBruts.Add(@{
      accountId        = $accId
      displayName      = $displayName
      email            = $email
      accountType      = $accType
      status           = $accStatus
      domaines         = $domainesStr
      onJiradot        = $onJiradot
      onMutexfr        = $onMutexfr
      dateCreation     = $dateCreation
      dateCreationFr   = Format-DateFr $dateCreation
      dateInvitation   = $dateInvitation
      dateInvitationFr = Format-DateFr $dateInvitation
      dateReferenceFr  = Format-DateFr $dateReference
      lastActive       = $lastActive
      lastActiveFr     = Format-DateFr $lastActive
      products         = $productList
      groupes          = ""
      action           = ""
    }) | Out-Null

    Start-Sleep -Milliseconds 100
  }

  # Progression
  Log ("  Page {0} : {1} traites, {2} pre-candidats, {3} avec produit (API)" -f ([int]($cTotal/100)), $cTotal, $cSansProduitAPI, $cAvecProduitAPI)

  # Pagination
  $cursor = $null
  if ($json.links -and $json.links.next) {
    $nextUrl = [string]$json.links.next
    if ($nextUrl -match "cursor=([^&]+)") { $cursor = $matches[1] }
  }
  if (-not $cursor) { break }
  Start-Sleep -Milliseconds 200
}

Log "=== Boucle principale terminee ==="
Log ("  Total scannes            : {0}" -f $cTotal)
Log ("  Skip - deja inactifs     : {0}" -f $cSkipInactive)
Log ("  Skip - comptes app       : {0}" -f $cSkipApp)
Log ("  Skip - hors Jiradot/mfr  : {0}" -f $cSkipHorsDomaine)
Log ("  Skip - email hors domaine: {0}" -f $cSkipHorsEmail)
Log ("  Skip - inactif cote site : {0}" -f $cSkipSiteInactif)
Log ("  Avec produit (API)       : {0}" -f $cAvecProduitAPI)
Log ("  Trop recents (< {0}m)    : {1}" -f $SeuilMois, $cTropRecent)
Log ("  Erreurs API profil       : {0}" -f $cErreurs)
Log ("  Pre-candidats (API=vide) : {0}" -f $candidatsBruts.Count)

# ============================================================
# ETAPE 4 : GROUPES + JIRA (applicationrole) + CONFLUENCE
#
# Pour chaque pre-candidat :
#   1. Fetch ses groupes (Jiradot + mutexfr)
#   2. Si un groupe est dans applicationrole → a un produit Jira → EXCLURE
#   3. Sinon, verifier l acces Confluence via API directe → EXCLURE si 200
#   4. Si uniquement groupes customer → EXCLURE
#   5. Sinon → CANDIDAT CONFIRME
# ============================================================

Log "=== ETAPE 4 : Verification groupes + acces Jira (applicationrole) + Confluence ==="

$candidats = New-Object System.Collections.Generic.List[object]
$cGroupFetched      = 0
$cJiraProductGroup  = 0
$cConfluenceAccess  = 0
$cCustomerOnly      = 0
$cPreCandidats      = $candidatsBruts.Count

foreach ($c in $candidatsBruts) {
  $cGroupFetched++

  if ($cGroupFetched % 10 -eq 0) {
    Write-Progress -Activity "Verification produit (groupes + Confluence)" `
      -Status ("{0}/{1} — Jira:{2} Confl:{3} Cust:{4}" -f $cGroupFetched, $cPreCandidats, $cJiraProductGroup, $cConfluenceAccess, $cCustomerOnly) `
      -PercentComplete ([int](100*$cGroupFetched/$cPreCandidats))
  }

  $groupNames = New-Object System.Collections.Generic.List[string]

  # Fetch groupes sur Jiradot
  if ($c.onJiradot) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteJiradot.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteJiradot.Headers
    if ($resp.ok) {
      $groups = $resp.content | ConvertFrom-Json
      foreach ($g in $groups) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) {
          $groupNames.Add($gName) | Out-Null
        }
      }
    }
  }

  # Fetch groupes sur mutexfr
  if ($c.onMutexfr) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteMutexfr.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteMutexfr.Headers
    if ($resp.ok) {
      $groups = $resp.content | ConvertFrom-Json
      foreach ($g in $groups) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) {
          $groupNames.Add($gName) | Out-Null
        }
      }
    }
  }

  $c["groupes"] = if ($groupNames.Count -gt 0) { ($groupNames | Sort-Object) -join " | " } else { "" }

  # --- FILTRE 8a : Groupe Jira produit (applicationrole) → exclure ---
  if (Test-HasJiraProductGroup $c.groupes) {
    $cJiraProductGroup++
    Log ("  JIRA (groupe) : {0} ({1})" -f $c.displayName, $c.email) "DEBUG"
    continue
  }

  # --- FILTRE 8b : Acces Confluence direct → exclure ---
  $hasConfluence = $false

  if ($c.onJiradot) {
    if (Test-ConfluenceAccess -Site $siteJiradot -AccountId $c.accountId) {
      $hasConfluence = $true
    }
  }

  if (-not $hasConfluence -and $c.onMutexfr) {
    if (Test-ConfluenceAccess -Site $siteMutexfr -AccountId $c.accountId) {
      $hasConfluence = $true
    }
  }

  if ($hasConfluence) {
    $cConfluenceAccess++
    Log ("  CONFLUENCE : {0} ({1})" -f $c.displayName, $c.email) "DEBUG"
    continue
  }

  # --- FILTRE 8c : Customer-only → exclure ---
  if (Test-IsCustomerOnly $c.groupes) {
    $cCustomerOnly++
    Log ("  CUSTOMER : {0} ({1})" -f $c.displayName, $c.email) "DEBUG"
    continue
  }

  # === CANDIDAT CONFIRME ===
  $candidats.Add($c) | Out-Null

  Start-Sleep -Milliseconds 100
}

Write-Progress -Activity "Verification produit" -Completed

$cCandidats = $candidats.Count

Log "=== Verification terminee ==="
Log ("  Pre-candidats (API sans produit) : {0}" -f $cPreCandidats)
Log ("  Exclus - Jira (applicationrole)  : {0}" -f $cJiraProductGroup)
Log ("  Exclus - Confluence (API directe): {0}" -f $cConfluenceAccess)
Log ("  Exclus - customer JSM only       : {0}" -f $cCustomerOnly)
Log ("  CANDIDATS CONFIRMES              : {0}" -f $cCandidats)

# ============================================================
# ETAPE 5 : EXPORT CSV
# ============================================================

Log "=== ETAPE 5 : Export CSV ==="

$csvColumns = @(
  "AccountId","DisplayName","Email","AccountType","Status",
  "Domaines","Groupes","DateCreation","DateInvitation","DerniereActivite",
  "Produits","Action"
)
$csvHeader = ($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
Write-CsvHeader $csvHeader

foreach ($c in ($candidats | Sort-Object { $_.displayName })) {
  $row = [ordered]@{
    AccountId        = $c.accountId
    DisplayName      = $c.displayName
    Email            = $c.email
    AccountType      = $c.accountType
    Status           = $c.status
    Domaines         = $c.domaines
    Groupes          = $c.groupes
    DateCreation     = $c.dateCreationFr
    DateInvitation   = $c.dateInvitationFr
    DerniereActivite = $c.lastActiveFr
    Produits         = $c.products
    Action           = "Candidat"
  }
  $line = ($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
  Write-CsvLine $line
}

Log "  CSV candidats -> $csvFile"
Log "  CSV groupes-roles -> $csvRolesFile"

# ============================================================
# ETAPE 6 : AFFICHAGE / DESACTIVATION
# ============================================================

if ($candidats.Count -eq 0) {
  Log "Aucun compte candidat a la desactivation."
  Write-Host ""
  Write-Host "  Aucun compte reellement sans produit de plus de $SeuilMois mois trouve." -ForegroundColor Green
  Write-Host ""
} else {
  Write-Host ""
  Write-Host "========================================" -ForegroundColor Cyan
  Write-Host ("  {0} COMPTES SANS PRODUIT CONFIRMES (> {1} mois)" -f $candidats.Count, $SeuilMois) -ForegroundColor Cyan
  Write-Host "  (hors Jira, hors Confluence, hors customers JSM)" -ForegroundColor DarkGray
  Write-Host "  Domaines : Jiradot / mutexfr" -ForegroundColor Cyan
  Write-Host "  Emails   : $($allowedDomains -join ', ')" -ForegroundColor Cyan
  Write-Host ("  Exclus Jira (applicationrole)   : {0}" -f $cJiraProductGroup) -ForegroundColor DarkGray
  Write-Host ("  Exclus Confluence (API directe) : {0}" -f $cConfluenceAccess) -ForegroundColor DarkGray
  Write-Host ("  Exclus customers JSM            : {0}" -f $cCustomerOnly) -ForegroundColor DarkGray
  Write-Host "========================================" -ForegroundColor Cyan
  Write-Host ""

  $idx = 0
  foreach ($c in ($candidats | Sort-Object { $_.displayName })) {
    $idx++
    $dateCr  = if ($c.dateCreationFr) { $c.dateCreationFr } else { "-" }
    $dateInv = if ($c.dateInvitationFr) { $c.dateInvitationFr } else { "-" }
    $lastAct = if ($c.lastActiveFr) { $c.lastActiveFr } else { "jamais" }
    $grpDisp = if ($c.groupes) { $c.groupes } else { "(aucun groupe)" }

    Write-Host ("{0,4}. {1}" -f $idx, $c.displayName) -ForegroundColor Yellow
    Write-Host ("      Email      : {0}" -f $c.email) -ForegroundColor DarkGray
    Write-Host ("      Domaines   : {0}" -f $c.domaines) -ForegroundColor DarkGray
    Write-Host ("      Creation   : {0}  |  Invitation : {1}" -f $dateCr, $dateInv) -ForegroundColor DarkGray
    Write-Host ("      Dern. act. : {0}" -f $lastAct) -ForegroundColor DarkGray
    Write-Host ("      Groupes    : {0}" -f $grpDisp) -ForegroundColor DarkCyan
    Write-Host ""
  }

  if ($Mode -eq "Dry") {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  MODE DRY : aucune desactivation" -ForegroundColor Green
    Write-Host "  Relancez avec -Mode Execute pour traiter" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Log "Mode Dry : aucune action effectuee."
  }
  else {
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  MODE EXECUTE : desactivation au cas par cas" -ForegroundColor Red
    Write-Host "  O=desactiver  N=ignorer  T=desactiver tous les restants  Q=quitter" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""

    $cDesactive = 0; $cIgnore = 0; $cErrDeact = 0; $autoAll = $false

    foreach ($c in ($candidats | Sort-Object { $_.displayName })) {
      $dateCr  = if ($c.dateCreationFr) { $c.dateCreationFr } else { "inconnue" }
      $dateInv = if ($c.dateInvitationFr) { $c.dateInvitationFr } else { "inconnue" }
      $lastAct = if ($c.lastActiveFr) { $c.lastActiveFr } else { "jamais" }
      $grpDisp = if ($c.groupes) { $c.groupes } else { "(aucun)" }

      Write-Host "----------------------------------------" -ForegroundColor DarkGray
      Write-Host ("  Nom         : {0}" -f $c.displayName) -ForegroundColor White
      Write-Host ("  Email       : {0}" -f $c.email) -ForegroundColor White
      Write-Host ("  Domaines    : {0}" -f $c.domaines) -ForegroundColor White
      Write-Host ("  Groupes     : {0}" -f $grpDisp) -ForegroundColor Cyan
      Write-Host ("  Type        : {0}" -f $c.accountType) -ForegroundColor White
      Write-Host ("  Creation    : {0}" -f $dateCr) -ForegroundColor White
      Write-Host ("  Invitation  : {0}" -f $dateInv) -ForegroundColor White
      Write-Host ("  Dern. activ.: {0}" -f $lastAct) -ForegroundColor White
      Write-Host ("  ID          : {0}" -f $c.accountId) -ForegroundColor DarkGray
      Write-Host ""

      $doIt = $false

      if ($autoAll) {
        $doIt = $true
        Write-Host "  => (auto) Desactivation..." -ForegroundColor DarkYellow
      } else {
        $choice = Read-Host "  Desactiver ? (O/N/T=tous/Q=quitter)"

        if ($choice -match "^[Qq]") {
          Log "Arret demande par l'utilisateur."
          Write-Host "  Arret." -ForegroundColor Yellow
          break
        }
        elseif ($choice -match "^[Tt]") {
          Write-Host ""
          $confirmAll = Read-Host "  Confirmer la desactivation de TOUS les comptes restants ? (OUI)"
          if ($confirmAll -eq "OUI") {
            $autoAll = $true
            $doIt = $true
            Log "Mode auto-all active par l'utilisateur."
          } else {
            Write-Host "  Auto-all annule, on continue au cas par cas." -ForegroundColor Yellow
            $choice = Read-Host "  Desactiver CE compte ? (O/N)"
            if ($choice -match "^[Oo]") { $doIt = $true }
          }
        }
        elseif ($choice -match "^[Oo]") {
          $doIt = $true
        }
      }

      if ($doIt) {
        $deactUrl = "https://api.atlassian.com/users/$($c.accountId)/manage/lifecycle/disable"
        $deactResp = Invoke-ApiCall -Method "POST" -Url $deactUrl -Headers $orgHeaders

        if ($deactResp.ok -or $deactResp.status -eq 204 -or $deactResp.status -eq 200) {
          Log ("  DESACTIVE : {0} ({1}) [{2}]" -f $c.displayName, $c.email, $c.domaines)
          Write-Host "  => DESACTIVE" -ForegroundColor Green
          $c.action = "Desactive"
          $cDesactive++
        } else {
          Log ("  ERREUR desactivation {0} : status={1} {2}" -f $c.displayName, $deactResp.status, $deactResp.error) "ERROR"
          Write-Host ("  => ERREUR (status={0})" -f $deactResp.status) -ForegroundColor Red
          $c.action = "Erreur"
          $cErrDeact++
        }
      } else {
        Log ("  IGNORE : {0} ({1})" -f $c.displayName, $c.email)
        Write-Host "  => Ignore" -ForegroundColor DarkGray
        $c.action = "Ignore"
        $cIgnore++
      }

      Start-Sleep -Milliseconds 300
    }

    # Reecrire le CSV avec les actions finales
    Write-CsvHeader $csvHeader
    foreach ($c in ($candidats | Sort-Object { $_.displayName })) {
      $action = if ($c.action) { $c.action } else { "Non traite" }
      $row = [ordered]@{
        AccountId        = $c.accountId
        DisplayName      = $c.displayName
        Email            = $c.email
        AccountType      = $c.accountType
        Status           = $c.status
        Domaines         = $c.domaines
        Groupes          = $c.groupes
        DateCreation     = $c.dateCreationFr
        DateInvitation   = $c.dateInvitationFr
        DerniereActivite = $c.lastActiveFr
        Produits         = $c.products
        Action           = $action
      }
      $line = ($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
      Write-CsvLine $line
    }

    Log ("Desactivation terminee : {0} desactives, {1} ignores, {2} erreurs" -f $cDesactive, $cIgnore, $cErrDeact)
  }
}

# ============================================================
# RESUME
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Total scannes (org)             : {0}" -f $cTotal)
Write-Host ("  Skip - deja inactifs            : {0}" -f $cSkipInactive) -ForegroundColor DarkGray
Write-Host ("  Skip - comptes app              : {0}" -f $cSkipApp) -ForegroundColor DarkGray
Write-Host ("  Skip - hors Jiradot/mutexfr     : {0}" -f $cSkipHorsDomaine) -ForegroundColor DarkGray
Write-Host ("  Skip - email hors domaine       : {0}" -f $cSkipHorsEmail) -ForegroundColor DarkGray
Write-Host ("  Skip - inactif cote site        : {0}" -f $cSkipSiteInactif) -ForegroundColor DarkGray
Write-Host ("  Avec produit (API)              : {0}" -f $cAvecProduitAPI) -ForegroundColor Green
Write-Host ("  Sans produit (API), trop recents: {0}" -f $cTropRecent) -ForegroundColor DarkGray
Write-Host ("  Pre-candidats (API sans produit): {0}" -f $cPreCandidats) -ForegroundColor DarkGray
Write-Host ("  Exclus - Jira (applicationrole) : {0}" -f $cJiraProductGroup) -ForegroundColor Green
Write-Host ("  Exclus - Confluence (API)       : {0}" -f $cConfluenceAccess) -ForegroundColor Green
Write-Host ("  Exclus - customers JSM only     : {0}" -f $cCustomerOnly) -ForegroundColor DarkGray
Write-Host ("  CANDIDATS CONFIRMES             : {0}" -f $cCandidats) -ForegroundColor $(if ($cCandidats -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Erreurs API                     : {0}" -f $cErreurs)
Write-Host ""
Write-Host ("  Seuil creation/invitation       : avant le {0}" -f $dateSeuilStr)
Write-Host ("  lastActive                      : informatif uniquement") -ForegroundColor DarkGray
Write-Host ("  Emails autorises                : {0}" -f ($allowedDomains -join ", "))
Write-Host ("  Groupes Jira produit            : {0} (via applicationrole)" -f $allProductGroupNames.Count)
Write-Host ("  Confluence                      : verification directe (API wiki)") -ForegroundColor DarkGray
Write-Host ("  Patterns groupes customer       : {0}" -f ($customerGroupPatterns -join ", ")) -ForegroundColor DarkGray
Write-Host ("  Mode                            : {0}" -f $Mode)
Write-Host ("  Duree                           : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ""
Write-Host "  Fichiers generes :" -ForegroundColor White
Write-Host ("    CSV candidats    : {0}" -f $csvFile)
Write-Host ("    CSV groupes/roles: {0}" -f $csvRolesFile)
Write-Host ("    LOG              : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."