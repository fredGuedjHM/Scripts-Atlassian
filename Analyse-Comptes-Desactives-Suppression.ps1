<#
.SYNOPSIS
  Analyse-Comptes-Desactives-Suppression.ps1
  Identifie les comptes desactives ou suspendus depuis plus de 3 mois
  en vue de leur suppression definitive.

.DESCRIPTION
  Ce script :
  1. Pre-charge les utilisateurs des sites Jiradot et mutexfr
  2. Parcourt les utilisateurs manages de l organisation via l Admin API
  3. Identifie :
     - Les comptes DESACTIVES (inactive) depuis plus de N mois
     - Les comptes SUSPENDUS (suspended) sur LES DEUX sites depuis plus de N mois
  4. Pour les comptes DSIM/DSIT : interroge Jira Assets (CMDB)
     pour recuperer Date Sortie, Statut et Motif Sortie.
     La Date Sortie remplace la date de reference lifecycle pour le calcul du seuil.
  5. Recupere les groupes d habilitation
  6. Exporte un CSV consolide + un CSV par site

  CRITERES D INCLUSION :
    A) account_status = "inactive"
       ET (date de sortie Asset OU date de reference lifecycle) > N mois
    B) account_status = "suspended"
       ET suspendu sur les 2 sites
       ET (date de sortie Asset OU date de reference lifecycle) > N mois

  ENRICHISSEMENT ASSET (comptes DSIM/DSIT) :
    - Endpoint : /gateway/api/jsm/assets/workspace/{wId}/v1/object/aql
    - Types : "Employe", "Prestataire" (schema Referentiel personne, key=RP)
    - Recherche par "Compte Jira" (attribut Utilisateur = accountId)
    - Attributs par ID : 408=Statut, 380=Date Sortie, 381=Motif Sortie, 379=Date Entree
    - Si Date Sortie est renseignee, elle remplace lastActive/dateCreation
      pour le calcul du seuil de 3 mois

  MODE : Dry-run uniquement (aucune suppression)

  CREDENTIALS :
    secrets\org-admin.xml    => OrgId + API Key
    secrets\site-admin.xml   => Jiradot
    secrets\site-mutexfr.xml => mutexfr

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 1.2 — Juillet 2026
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Analyse-Comptes-Desactives-Suppression.ps1
  .\Analyse-Comptes-Desactives-Suppression.ps1 -SeuilMois 6
#>

[CmdletBinding()]
param(
  [int] $SeuilMois = 3,

  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile       = Join-Path $ExportsDir ("ComptesDesactives-Suppression_{0}.log" -f $ts)
$csvConsolide  = Join-Path $ExportsDir ("ComptesDesactives-Suppression_TOUS_{0}.csv" -f $ts)
$csvJiradot    = Join-Path $ExportsDir ("ComptesDesactives-Suppression_Jiradot_{0}.csv" -f $ts)
$csvMutexfr    = Join-Path $ExportsDir ("ComptesDesactives-Suppression_mutexfr_{0}.csv" -f $ts)
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
    $dt = [DateTimeOffset]::Parse($isoDate, [System.Globalization.CultureInfo]::InvariantCulture)
    return $dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
  } catch { return $isoDate }
}

function Get-AncienneteJours([string]$isoDate) {
  if (-not $isoDate) { return -1 }
  try {
    $dt = [DateTimeOffset]::Parse($isoDate, [System.Globalization.CultureInfo]::InvariantCulture)
    return [int]((Get-Date) - $dt.DateTime).TotalDays
  } catch { return -1 }
}

function ConvertTo-DateFromAsset([string]$assetDate) {
    <#
    Parse les dates Assets qui peuvent etre en format :
      - "31/May/26" (dd/MMM/yy)
      - "17/Feb/25 9:48 AM" (dd/MMM/yy h:mm tt)
      - "2026-05-31" (ISO)
      - "31/05/2026" (dd/MM/yyyy)
    Retourne une date ISO ou la valeur originale si echec.
    #>
    if ([string]::IsNullOrWhiteSpace($assetDate)) { return "" }

    # Essayer le parsing standard
    try {
        $dt = [DateTimeOffset]::Parse($assetDate, [System.Globalization.CultureInfo]::InvariantCulture)
        return $dt.ToString("yyyy-MM-dd'T'HH:mm:sszzz")
    } catch {}

    # Essayer les formats Assets specifiques
    $formats = @(
        "dd/MMM/yy",
        "dd/MMM/yy h:mm tt",
        "dd/MMM/yyyy",
        "dd/MMM/yyyy h:mm tt",
        "dd/MMM/yy H:mm",
        "dd/MMM/yyyy H:mm"
    )
    foreach ($fmt in $formats) {
        try {
            $dt = [DateTime]::ParseExact($assetDate, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            return $dt.ToString("yyyy-MM-dd'T'00:00:00+02:00")
        } catch {}
    }

    # Retourner tel quel si aucun format ne matche
    return $assetDate
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
# DETECTION GROUPE DSIM / DSIT
# ============================================================

function Test-IsDsimGroup([string]$groupesStr) {
    if ([string]::IsNullOrWhiteSpace($groupesStr)) { return $false }
    $lower = $groupesStr.ToLower()
    return ($lower.Contains("dsim") -or $lower.Contains("dsit"))
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

function Import-SiteCredentials([string]$FilePath, [string]$SiteName) {
  if (-not (Test-Path $FilePath)) {
    Log "Fichier $FilePath introuvable, creation interactive..." "WARN"
    [System.Windows.Forms.MessageBox]::Show(
      ("Credentials pour {0} manquants.`n`nVous allez fournir :`n  1. URL du site`n  2. Email admin`n  3. API Token" -f $SiteName),
      "Credentials $SiteName", [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl = Read-Host "URL du site (ex: $SiteName.atlassian.net)"
    $inputUrl = $inputUrl -replace "^https?://", "" -replace "/.*$", "" -replace "/$", ""
    $adminEmail = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $FilePath
    Log "Credentials sauvegardes dans $FilePath"
  }
  $data = Import-Clixml -Path $FilePath
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

function Get-SiteUsersWithStatus([hashtable]$Site) {
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
# JIRA ASSETS — RECHERCHE PAR COMPTE JIRA (accountId)
# Endpoint : /gateway/api/jsm/assets/workspace/{wId}/v1/object/aql
# Types    : "Employé", "Prestataire" (schema Référentiel personne, key=RP)
# Attributs par ID : 408=Statut, 380=Date Sortie, 381=Motif Sortie, 379=Date Entree
# ============================================================

function Get-AssetsWorkspaceId([hashtable]$Site) {
    $url = "{0}/rest/servicedeskapi/assets/workspace" -f $Site.BaseUrl
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    if ($resp.ok) {
        $json = $resp.content | ConvertFrom-Json
        if ($json.values -and $json.values.Count -gt 0) {
            return [string]$json.values[0].workspaceId
        }
    }
    return $null
}

function Search-AssetByAccountId([hashtable]$Site, [string]$WorkspaceId, [string]$AccountId) {
    <#
    Recherche un objet Employe/Prestataire dans Assets par "Compte Jira" (accountId).
    L endpoint Gateway ne retourne PAS objectTypeAttribute.name,
    seulement objectTypeAttributeId. On identifie les attributs par leur ID :
      408 = Statut (Etat : Actif/Inactif)
      380 = Date Sortie
      381 = Motif Sortie
      379 = Date Entree
    #>
    if ([string]::IsNullOrWhiteSpace($AccountId) -or [string]::IsNullOrWhiteSpace($WorkspaceId)) {
        return $null
    }

    $url = "{0}/gateway/api/jsm/assets/workspace/{1}/v1/object/aql" -f $Site.BaseUrl, $WorkspaceId

    $aql = 'objectType IN ("Employé", "Prestataire") AND "Compte Jira" = "{0}"' -f $AccountId

    $body = @{
        qlQuery = $aql
        maxResults = 1
        includeAttributes = $true
    } | ConvertTo-Json -Depth 5 -Compress

    $resp = Invoke-ApiCall -Method "POST" -Url $url -Headers $Site.Headers -Body $body

    if (-not $resp.ok) {
        Log ("  ASSET : Erreur HTTP {0} pour {1}" -f $resp.status, $AccountId) "WARN"
        return $null
    }

    $json = $resp.content | ConvertFrom-Json

    $objects = $null
    if ($json.values) { $objects = $json.values }
    elseif ($json.objectEntries) { $objects = $json.objectEntries }
    elseif ($json -is [array]) { $objects = $json }

    if (-not $objects -or ($objects | Measure-Object).Count -eq 0) {
        return $null
    }

    $asset = $objects[0]
    $result = @{
        Statut      = ""
        DateSortie  = ""
        MotifSortie = ""
        DateEntree  = ""
        ObjectType  = ""
        ObjectKey   = ""
    }

    if ($asset.objectType -and $asset.objectType.name) {
        $result.ObjectType = [string]$asset.objectType.name
    }
    if ($asset.objectKey) {
        $result.ObjectKey = [string]$asset.objectKey
    }

    # Extraction par ID d attribut (Gateway ne fournit pas le nom)
    if ($asset.attributes) {
        foreach ($attr in $asset.attributes) {
            $attrId = ""
            if ($attr.objectTypeAttributeId) {
                $attrId = [string]$attr.objectTypeAttributeId
            } elseif ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.id) {
                $attrId = [string]$attr.objectTypeAttribute.id
            }

            $attrVal = ""
            if ($attr.objectAttributeValues -and $attr.objectAttributeValues.Count -gt 0) {
                $val = $attr.objectAttributeValues[0]
                if ($val.displayValue) { $attrVal = [string]$val.displayValue }
                elseif ($val.value) { $attrVal = [string]$val.value }
                elseif ($val.status -and $val.status.name) { $attrVal = [string]$val.status.name }
            }

            switch ($attrId) {
                "408" { $result.Statut = $attrVal }
                "380" { $result.DateSortie = $attrVal }
                "381" { $result.MotifSortie = $attrVal }
                "379" { $result.DateEntree = $attrVal }
            }
        }
    }

    return $result
}

# ============================================================
# CONFIGURATION
# ============================================================

$dateSeuil = (Get-Date).AddMonths(-$SeuilMois)
$dateSeuilStr = $dateSeuil.ToString("dd/MM/yyyy")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ANALYSE COMPTES DESACTIVES / SUSPENDUS" -ForegroundColor Cyan
Write-Host "  en vue de SUPPRESSION DEFINITIVE" -ForegroundColor Cyan
Write-Host ("  Seuil : > {0} mois (avant le {1})" -f $SeuilMois, $dateSeuilStr) -ForegroundColor Cyan
Write-Host "  Sites : Jiradot + mutexfr" -ForegroundColor Cyan
Write-Host "  Inclus : inactive (tous) + suspended (si sur les 2 sites)" -ForegroundColor Cyan
Write-Host "  Enrichissement : Assets CMDB (comptes DSIM/DSIT)" -ForegroundColor Cyan
Write-Host "  Mode : DRY-RUN (analyse uniquement)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  /!\ La suppression est IRREVERSIBLE." -ForegroundColor Red
Write-Host "  Ce script identifie les candidats. Aucune suppression n'est effectuee." -ForegroundColor Red
Write-Host ""

Log "================================================================"
Log "  ANALYSE COMPTES DESACTIVES / SUSPENDUS - EN VUE DE SUPPRESSION"
Log ("  Seuil : {0} mois (avant le {1})" -f $SeuilMois, $dateSeuilStr)
Log "  Sites : Jiradot + mutexfr"
Log "  Inclus : inactive (tous) + suspended (si sur les 2 sites)"
Log "  Enrichissement : Assets CMDB pour comptes DSIM/DSIT"
Log ("  Emails autorises : {0}" -f ($allowedDomains -join ", "))
Log "  Mode : DRY-RUN"
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

$siteJiradot = Import-SiteCredentials -FilePath (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
$siteMutexfr = Import-SiteCredentials -FilePath (Join-Path $SecretsDir "site-mutexfr.xml") -SiteName "mutexfr"

Log ("  OrgId    : {0}" -f $orgId)
Log ("  Jiradot  : {0}" -f $siteJiradot.BaseUrl)
Log ("  mutexfr  : {0}" -f $siteMutexfr.BaseUrl)

# ============================================================
# ETAPE 2 : PRE-CHARGEMENT DES UTILISATEURS PAR SITE
# ============================================================

Log "=== ETAPE 2 : Pre-chargement des utilisateurs par site (avec statut) ==="

$jiradotUsers = Get-SiteUsersWithStatus -Site $siteJiradot
$mutexfrUsers = Get-SiteUsersWithStatus -Site $siteMutexfr

# ============================================================
# ETAPE 2b : DECOUVERTE WORKSPACE ASSETS
# ============================================================

Log "=== ETAPE 2b : Decouverte du workspace Jira Assets ==="

$assetsWorkspaceId = Get-AssetsWorkspaceId -Site $siteJiradot

if ($assetsWorkspaceId) {
    Log ("  Workspace Assets Jiradot : {0}" -f $assetsWorkspaceId)
} else {
    Log "  WARN : Workspace Assets non trouve — enrichissement CMDB desactive" "WARN"
}

# ============================================================
# ETAPE 3 : IDENTIFICATION DES COMPTES DESACTIVES / SUSPENDUS
# ============================================================

Log "=== ETAPE 3 : Identification des comptes desactives/suspendus ==="

$comptesDesactives = New-Object System.Collections.Generic.List[object]
$cursor = $null
$startTime = Get-Date

$cTotal              = 0
$cSkipActif          = 0
$cSkipClosed         = 0
$cSkipApp            = 0
$cSkipHorsDomaine    = 0
$cSkipHorsEmail      = 0
$cSkipSuspPartiel    = 0
$cRetenuInactive     = 0
$cRetenuSuspended    = 0
$cErreurs            = 0

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

    if ($accStatus -eq "active") { $cSkipActif++; continue }
    if ($accStatus -eq "closed") { $cSkipClosed++; continue }
    if ($accStatus -ne "inactive" -and $accStatus -ne "suspended") { $cSkipActif++; continue }
    if ($accType -eq "app") { $cSkipApp++; continue }

    $onJiradot = $jiradotUsers.ContainsKey($accId)
    $onMutexfr = $mutexfrUsers.ContainsKey($accId)
    if (-not $onJiradot -and -not $onMutexfr) { $cSkipHorsDomaine++; continue }

    if ($email -and -not (Test-EmailDomainAllowed $email)) { $cSkipHorsEmail++; continue }

    # Filtre suspended : doit etre inactif sur tous les sites
    if ($accStatus -eq "suspended") {
      $suspendedEverywhere = $true
      if ($onJiradot -and $jiradotUsers[$accId].active -eq $true) { $suspendedEverywhere = $false }
      if ($onMutexfr -and $mutexfrUsers[$accId].active -eq $true) { $suspendedEverywhere = $false }
      if (-not $suspendedEverywhere) { $cSkipSuspPartiel++; continue }
    }

    # Fetch profil (dates)
    $profileUrl = "https://api.atlassian.com/users/$accId/manage/profile"
    $profResp = Invoke-ApiCall -Method "GET" -Url $profileUrl -Headers $orgHeaders

    $dateCreation = ""
    $dateInvitation = ""

    if ($profResp.ok) {
      $profData = $profResp.content | ConvertFrom-Json
      if (-not $email) {
        if ($profData.account -and $profData.account.email) { $email = [string]$profData.account.email }
        elseif ($profData.email) { $email = [string]$profData.email }
        if ($email -and -not (Test-EmailDomainAllowed $email)) { $cSkipHorsEmail++; continue }
      }
      if ($profData.account -and $profData.account.created) { $dateCreation = [string]$profData.account.created }
      elseif ($profData.created) { $dateCreation = [string]$profData.created }
      if ($profData.account -and $profData.account.invited) { $dateInvitation = [string]$profData.account.invited }
      elseif ($profData.invited) { $dateInvitation = [string]$profData.invited }
    } else {
      if ($profResp.status -ne 404) {
        Log ("  Erreur profil {0} ({1}) : status={2}" -f $displayName, $accId, $profResp.status) "WARN"
        $cErreurs++
      }
    }

    # Date de reference lifecycle (sera potentiellement remplacee par Date Sortie Asset)
    $dateReference = ""
    if ($lastActive) { $dateReference = $lastActive }
    elseif ($dateInvitation) { $dateReference = $dateInvitation }
    elseif ($dateCreation) { $dateReference = $dateCreation }

    $domaines = @()
    if ($onJiradot) { $domaines += "Jiradot" }
    if ($onMutexfr) { $domaines += "mutexfr" }
    $domainesStr = $domaines -join " + "

    $comptesDesactives.Add(@{
      accountId         = $accId
      displayName       = $displayName
      email             = $email
      accountType       = $accType
      status            = $accStatus
      domaines          = $domainesStr
      onJiradot         = $onJiradot
      onMutexfr         = $onMutexfr
      dateCreation      = $dateCreation
      dateCreationFr    = Format-DateFr $dateCreation
      dateInvitation    = $dateInvitation
      dateInvitationFr  = Format-DateFr $dateInvitation
      lastActive        = $lastActive
      lastActiveFr      = if ($lastActive) { Format-DateFr $lastActive } else { "(jamais)" }
      dateReference     = $dateReference
      dateReferenceFr   = Format-DateFr $dateReference
      ancienneteJours   = -1
      ancienneteMois    = -1
      groupes           = ""
      isDsim            = $false
      statutAsset       = ""
      dateSortieAsset   = ""
      dateSortieAssetFr = ""
      motifSortie       = ""
    }) | Out-Null

    if ($accStatus -eq "inactive") { $cRetenuInactive++ }
    else { $cRetenuSuspended++ }

    Start-Sleep -Milliseconds 100
  }

  $cRetenuTotal = $cRetenuInactive + $cRetenuSuspended
  Log ("  Page {0} : {1} traites, {2} retenus" -f ([int]($cTotal/100)), $cTotal, $cRetenuTotal)

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
Log ("  Skip - actifs            : {0}" -f $cSkipActif)
Log ("  Skip - closed            : {0}" -f $cSkipClosed)
Log ("  Skip - comptes app       : {0}" -f $cSkipApp)
Log ("  Skip - hors Jiradot/mfr  : {0}" -f $cSkipHorsDomaine)
Log ("  Skip - email hors domaine: {0}" -f $cSkipHorsEmail)
Log ("  Skip - suspended partiel : {0}" -f $cSkipSuspPartiel)
Log ("  Erreurs API              : {0}" -f $cErreurs)
Log ("  PRE-RETENUS (avant seuil): {0} (inactive:{1} suspended:{2})" -f $comptesDesactives.Count, $cRetenuInactive, $cRetenuSuspended)

# ============================================================
# ETAPE 4 : GROUPES + ENRICHISSEMENT ASSET (DSIM/DSIT)
#            + APPLICATION DU SEUIL
# ============================================================

Log "=== ETAPE 4 : Groupes + Assets CMDB + application seuil ==="

$cFetched = 0
$cCount = $comptesDesactives.Count
$cDsim = 0
$cAssetTrouve = 0
$cAssetNonTrouve = 0
$cFiltreParSeuil = 0

$comptesFinaux = New-Object System.Collections.Generic.List[object]

foreach ($c in $comptesDesactives) {
  $cFetched++

  if ($cFetched % 10 -eq 0) {
    Write-Progress -Activity "Groupes + Assets + seuil" `
      -Status ("{0}/{1}" -f $cFetched, $cCount) `
      -PercentComplete ([int](100*$cFetched/$cCount))
  }

  # --- Recuperation des groupes ---
  $groupNames = New-Object System.Collections.Generic.List[string]

  if ($c.onJiradot) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteJiradot.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteJiradot.Headers
    if ($resp.ok) {
      $groups = $resp.content | ConvertFrom-Json
      foreach ($g in $groups) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) { $groupNames.Add($gName) | Out-Null }
      }
    }
  }

  if ($c.onMutexfr) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteMutexfr.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteMutexfr.Headers
    if ($resp.ok) {
      $groups = $resp.content | ConvertFrom-Json
      foreach ($g in $groups) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) { $groupNames.Add($gName) | Out-Null }
      }
    }
  }

  $c["groupes"] = if ($groupNames.Count -gt 0) { ($groupNames | Sort-Object) -join " | " } else { "" }
  $c["isDsim"] = Test-IsDsimGroup $c.groupes

  # --- Enrichissement Asset pour les comptes DSIM/DSIT ---
  if ($c.isDsim -and $assetsWorkspaceId) {
    $cDsim++

    $assetData = Search-AssetByAccountId -Site $siteJiradot -WorkspaceId $assetsWorkspaceId -AccountId $c.accountId

    if ($assetData) {
      $cAssetTrouve++
      $c["statutAsset"] = $assetData.Statut

      # Convertir la date Asset en format parseable
      $dateSortieIso = ConvertTo-DateFromAsset $assetData.DateSortie
      $c["dateSortieAsset"] = $dateSortieIso
      $c["dateSortieAssetFr"] = Format-DateFr $dateSortieIso
      $c["motifSortie"] = $assetData.MotifSortie

      # Si Date Sortie existe, elle REMPLACE la date de reference lifecycle
      if ($dateSortieIso) {
        $c["dateReference"] = $dateSortieIso
        $c["dateReferenceFr"] = Format-DateFr $dateSortieIso
      }

      Log ("  ASSET : {0} -> Statut={1}, Sortie={2}, Motif={3}" -f $c.displayName, $assetData.Statut, (Format-DateFr $dateSortieIso), $assetData.MotifSortie) "DEBUG"
    } else {
      $cAssetNonTrouve++
      $c["statutAsset"] = "Non trouvé"
      Log ("  ASSET : {0} -> Non trouvé dans CMDB" -f $c.displayName) "DEBUG"
    }
  } elseif ($c.isDsim -and -not $assetsWorkspaceId) {
    $c["statutAsset"] = "(Assets indisponible)"
    $cDsim++
  }

  # --- Application du seuil sur la date de reference (finale) ---
  $dateRef = $c.dateReference
  $isOldEnough = $false

  if ($dateRef) {
    try {
      $dtRef = [DateTimeOffset]::Parse($dateRef, [System.Globalization.CultureInfo]::InvariantCulture)
      if ($dtRef.DateTime -lt $dateSeuil) { $isOldEnough = $true }
    } catch {}
  } else {
    # Pas de date du tout -> inclure par prudence
    $isOldEnough = $true
  }

  if (-not $isOldEnough) {
    $cFiltreParSeuil++
    continue
  }

  # Calculer l anciennete finale
  $ancienneteJours = Get-AncienneteJours $c.dateReference
  $c["ancienneteJours"] = $ancienneteJours
  $c["ancienneteMois"] = if ($ancienneteJours -gt 0) { [Math]::Round($ancienneteJours / 30.44, 1) } else { -1 }

  $comptesFinaux.Add($c) | Out-Null

  Start-Sleep -Milliseconds 100
}

Write-Progress -Activity "Groupes + Assets" -Completed

Log ("  Groupes recuperes pour      : {0} comptes" -f $cCount)
Log ("  Comptes DSIM/DSIT           : {0}" -f $cDsim)
Log ("  Asset trouve dans CMDB      : {0}" -f $cAssetTrouve)
Log ("  Asset non trouve            : {0}" -f $cAssetNonTrouve)
Log ("  Filtres par seuil (< {0}m)   : {1}" -f $SeuilMois, $cFiltreParSeuil)
Log ("  CANDIDATS FINAUX            : {0}" -f $comptesFinaux.Count)
# ============================================================
# ETAPE 5 : EXPORT CSV
# ============================================================

Log "=== ETAPE 5 : Export CSV ==="

$csvColumns = @(
  "AccountId","DisplayName","Email","AccountType","Status",
  "Domaines","Groupes","DateCreation","DateInvitation","DerniereActivite",
  "DateReference","AncienneteMois",
  "StatutAsset","DateSortieAsset","MotifSortie"
)

function Export-CsvDesactives([string]$FilePath, [System.Collections.Generic.List[object]]$Data) {
  $header = ($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine($header)

  foreach ($c in ($Data | Sort-Object { $_.ancienneteJours } -Descending)) {
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
      DateReference    = $c.dateReferenceFr
      AncienneteMois   = if ($c.ancienneteMois -gt 0) { $c.ancienneteMois.ToString("0.0") } else { "?" }
      StatutAsset      = $c.statutAsset
      DateSortieAsset  = $c.dateSortieAssetFr
      MotifSortie      = $c.motifSortie
    }
    $line = ($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
    [void]$sb.AppendLine($line)
  }

  [System.IO.File]::WriteAllText($FilePath, $sb.ToString(), $utf8Bom)
}

# CSV consolide
Export-CsvDesactives -FilePath $csvConsolide -Data $comptesFinaux
Log "  CSV consolide -> $csvConsolide"

# CSV Jiradot
$dataJiradot = New-Object System.Collections.Generic.List[object]
foreach ($c in $comptesFinaux) {
  if ($c.onJiradot) { $dataJiradot.Add($c) | Out-Null }
}
Export-CsvDesactives -FilePath $csvJiradot -Data $dataJiradot
Log ("  CSV Jiradot -> {0} ({1} comptes)" -f $csvJiradot, $dataJiradot.Count)

# CSV mutexfr
$dataMutexfr = New-Object System.Collections.Generic.List[object]
foreach ($c in $comptesFinaux) {
  if ($c.onMutexfr) { $dataMutexfr.Add($c) | Out-Null }
}
Export-CsvDesactives -FilePath $csvMutexfr -Data $dataMutexfr
Log ("  CSV mutexfr -> {0} ({1} comptes)" -f $csvMutexfr, $dataMutexfr.Count)

# ============================================================
# ETAPE 6 : AFFICHAGE
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  {0} COMPTES CANDIDATS A LA SUPPRESSION" -f $comptesFinaux.Count) -ForegroundColor Cyan
Write-Host ("  (desactives/suspendus > {0} mois)" -f $SeuilMois) -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Repartition par statut
$countInactive  = ($comptesFinaux | Where-Object { $_.status -eq "inactive" }).Count
$countSuspended = ($comptesFinaux | Where-Object { $_.status -eq "suspended" }).Count

Write-Host ("  Par statut :") -ForegroundColor White
Write-Host ("    inactive (desactives)         : {0}" -f $countInactive) -ForegroundColor Yellow
Write-Host ("    suspended (suspendus 2 sites) : {0}" -f $countSuspended) -ForegroundColor Yellow
Write-Host ""

Write-Host ("  Par site :") -ForegroundColor White
Write-Host ("    Jiradot : {0}" -f $dataJiradot.Count) -ForegroundColor Yellow
Write-Host ("    mutexfr : {0}" -f $dataMutexfr.Count) -ForegroundColor Yellow
Write-Host ""

# Repartition DSIM/DSIT avec Asset
$countDsimAvecAsset = ($comptesFinaux | Where-Object { $_.isDsim -and $_.statutAsset -and $_.statutAsset -ne "Non trouvé" -and $_.statutAsset -ne "(Assets indisponible)" }).Count
$countDsimSansAsset = ($comptesFinaux | Where-Object { $_.isDsim -and ($_.statutAsset -eq "Non trouvé") }).Count
$countNonDsim       = ($comptesFinaux | Where-Object { -not $_.isDsim }).Count

Write-Host ("  Enrichissement CMDB :") -ForegroundColor White
Write-Host ("    DSIM/DSIT avec fiche Asset    : {0}" -f $countDsimAvecAsset) -ForegroundColor Cyan
Write-Host ("    DSIM/DSIT sans fiche Asset    : {0}" -f $countDsimSansAsset) -ForegroundColor DarkYellow
Write-Host ("    Hors DSIM/DSIT (lifecycle)    : {0}" -f $countNonDsim) -ForegroundColor DarkGray
Write-Host ""

# Statistiques par tranche d anciennete
$tranche3_6   = ($comptesFinaux | Where-Object { $_.ancienneteMois -ge 3 -and $_.ancienneteMois -lt 6 }).Count
$tranche6_12  = ($comptesFinaux | Where-Object { $_.ancienneteMois -ge 6 -and $_.ancienneteMois -lt 12 }).Count
$tranche12_24 = ($comptesFinaux | Where-Object { $_.ancienneteMois -ge 12 -and $_.ancienneteMois -lt 24 }).Count
$tranche24p   = ($comptesFinaux | Where-Object { $_.ancienneteMois -ge 24 }).Count
$trancheIndet = ($comptesFinaux | Where-Object { $_.ancienneteMois -le 0 }).Count

Write-Host "  Repartition par anciennete :" -ForegroundColor White
Write-Host ("    3 - 6 mois    : {0}" -f $tranche3_6) -ForegroundColor DarkGray
Write-Host ("    6 - 12 mois   : {0}" -f $tranche6_12) -ForegroundColor DarkGray
Write-Host ("   12 - 24 mois   : {0}" -f $tranche12_24) -ForegroundColor Yellow
Write-Host ("   > 24 mois      : {0}" -f $tranche24p) -ForegroundColor Red
if ($trancheIndet -gt 0) {
  Write-Host ("    Indetermine   : {0}" -f $trancheIndet) -ForegroundColor DarkGray
}
Write-Host ""

# Affichage detaille (max 50)
$displayList = $comptesFinaux | Sort-Object { $_.ancienneteJours } -Descending
$displayMax = [Math]::Min(50, $displayList.Count)

$idx = 0
foreach ($c in $displayList) {
  $idx++
  if ($idx -gt $displayMax) {
    Write-Host ("      ... et {0} autres (voir CSV)" -f ($displayList.Count - $displayMax)) -ForegroundColor DarkGray
    break
  }

  $dateCr  = if ($c.dateCreationFr) { $c.dateCreationFr } else { "-" }
  $dateInv = if ($c.dateInvitationFr) { $c.dateInvitationFr } else { "-" }
  $lastAct = $c.lastActiveFr
  $grpDisp = if ($c.groupes) { $c.groupes } else { "(aucun groupe)" }
  $ancStr  = if ($c.ancienneteMois -gt 0) { "{0:0.0} mois" -f $c.ancienneteMois } else { "?" }
  $statDisp = if ($c.status -eq "suspended") { "[SUSPENDED]" } else { "[INACTIVE]" }

  $ancColor = "DarkGray"
  if ($c.ancienneteMois -ge 24) { $ancColor = "Red" }
  elseif ($c.ancienneteMois -ge 12) { $ancColor = "Yellow" }

  Write-Host ("{0,4}. {1} {2}" -f $idx, $c.displayName, $statDisp) -ForegroundColor $ancColor
  Write-Host ("      Email       : {0}" -f $c.email) -ForegroundColor DarkGray
  Write-Host ("      Domaines    : {0}" -f $c.domaines) -ForegroundColor DarkGray
  Write-Host ("      Creation    : {0}  |  Invitation : {1}" -f $dateCr, $dateInv) -ForegroundColor DarkGray
  Write-Host ("      Dern. act.  : {0}" -f $lastAct) -ForegroundColor DarkGray
  Write-Host ("      Anciennete  : {0}" -f $ancStr) -ForegroundColor $ancColor
  if ($c.isDsim) {
    $assetColor = if ($c.statutAsset -match "Inactif") { "Green" } elseif ($c.statutAsset -eq "Non trouvé") { "Red" } else { "Yellow" }
    Write-Host ("      CMDB Asset  : Statut={0} | Sortie={1} | Motif={2}" -f $c.statutAsset, $c.dateSortieAssetFr, $c.motifSortie) -ForegroundColor $assetColor
  }
  Write-Host ("      Groupes     : {0}" -f $grpDisp) -ForegroundColor DarkCyan
  Write-Host ""
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
Write-Host ("  Skip - actifs                   : {0}" -f $cSkipActif) -ForegroundColor DarkGray
Write-Host ("  Skip - closed                   : {0}" -f $cSkipClosed) -ForegroundColor DarkGray
Write-Host ("  Skip - comptes app              : {0}" -f $cSkipApp) -ForegroundColor DarkGray
Write-Host ("  Skip - hors Jiradot/mutexfr     : {0}" -f $cSkipHorsDomaine) -ForegroundColor DarkGray
Write-Host ("  Skip - email hors domaine       : {0}" -f $cSkipHorsEmail) -ForegroundColor DarkGray
Write-Host ("  Skip - suspended partiel        : {0}" -f $cSkipSuspPartiel) -ForegroundColor DarkGray
Write-Host ("  Skip - trop recents (< {0}m)     : {1}" -f $SeuilMois, $cFiltreParSeuil) -ForegroundColor DarkGray
Write-Host ("  RETENUS - inactive              : {0}" -f $countInactive) -ForegroundColor Yellow
Write-Host ("  RETENUS - suspended (2 sites)   : {0}" -f $countSuspended) -ForegroundColor Yellow
Write-Host ("  TOTAL CANDIDATS SUPPRESSION     : {0}" -f $comptesFinaux.Count) -ForegroundColor $(if ($comptesFinaux.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host ("    dont Jiradot                  : {0}" -f $dataJiradot.Count)
Write-Host ("    dont mutexfr                  : {0}" -f $dataMutexfr.Count)
Write-Host ("    dont DSIM/DSIT (Asset)        : {0}" -f $cDsim) -ForegroundColor Cyan
Write-Host ("      Asset trouve                : {0}" -f $cAssetTrouve) -ForegroundColor Green
Write-Host ("      Asset non trouve            : {0}" -f $cAssetNonTrouve) -ForegroundColor DarkYellow
Write-Host ("  Erreurs API                     : {0}" -f $cErreurs)
Write-Host ""
Write-Host ("  Seuil                           : > {0} mois (avant le {1})" -f $SeuilMois, $dateSeuilStr)
Write-Host ("  Emails autorises                : {0}" -f ($allowedDomains -join ", "))
Write-Host ("  Workspace Assets                : {0}" -f $(if ($assetsWorkspaceId) { $assetsWorkspaceId } else { "(indisponible)" }))
Write-Host ("  Endpoint Assets                 : /gateway/api/jsm/assets/...") -ForegroundColor DarkGray
Write-Host ("  Recherche Asset DSIM/DSIT       : Par Compte Jira (accountId)") -ForegroundColor DarkGray
Write-Host ("  Types CMDB                      : Employé, Prestataire") -ForegroundColor DarkGray
Write-Host ("  Attributs par ID                : 408=Statut, 380=DateSortie, 381=Motif, 379=DateEntree") -ForegroundColor DarkGray
Write-Host ("  Mode                            : DRY-RUN (analyse)") -ForegroundColor Green
Write-Host ("  Duree                           : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ""
Write-Host "  /!\ RAPPEL : La suppression est IRREVERSIBLE." -ForegroundColor Red
Write-Host "  Verifiez le CSV avant toute action future." -ForegroundColor Red
Write-Host ""
Write-Host "  Fichiers generes :" -ForegroundColor White
Write-Host ("    CSV consolide  : {0}" -f $csvConsolide)
Write-Host ("    CSV Jiradot    : {0}" -f $csvJiradot)
Write-Host ("    CSV mutexfr    : {0}" -f $csvMutexfr)
Write-Host ("    LOG            : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."