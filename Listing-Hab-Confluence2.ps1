<#
.SYNOPSIS
  Listing-Habilitations-Confluence.ps1 - Audit des habilitations Confluence
  avec dates, types d'utilisateurs et resolution email multi-sources.
.DESCRIPTION
  Ce script :
  1. Liste les espaces Confluence (depuis un fichier OU tous les globaux)
  2. Pour chaque espace, recupere permissions via API v2 + date de creation
  3. Pour chaque espace, recherche la date de derniere modification (CQL)
  4. Resout les groupes en membres (cache par groupe)
  5. Recupere l'email via 4 sources en cascade
  6. Identifie les Guests et Collaborateurs Externes
  7. Genere un CSV UTF-8 BOM (accents corrects dans Excel)

  Le CSV est ecrit AU FIL DE L'EAU (un espace a la fois).

  CREDENTIALS NECESSAIRES :
    - secrets\site-admin.xml : SiteUrl + Email + API Token
    - secrets\org-admin.xml  : OrgId + API Key

  MODES D'UTILISATION :
    - Avec -SpaceKeysFile : audite uniquement les espaces du fichier (rapide)
    - Sans parametre      : audite tous les espaces globaux (exclut les personnels)

  NOTE SUR "DATE DE DERNIER ACCES" :
    L'API audit Confluence Cloud ne trace PAS les consultations de pages.
    La colonne DateDerniereModif (derniere modification de contenu) est
    le meilleur proxy disponible via API.
    Pour les donnees de consultation, voir le dashboard Analytics :
        admin.atlassian.com > Analytics > Spaces.
.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 2.0 - Correction API v2 permissions + filtre espaces
  Compatibilite  : PowerShell 5.1+
.EXAMPLE
  .\Listing-Habilitations-Confluence.ps1 -SpaceKeysFile ".\spacekeys.txt"
  .\Listing-Habilitations-Confluence.ps1
  .\Listing-Habilitations-Confluence.ps1 -ThrottleMs 200
#>
[CmdletBinding()]
param(
  [int]    $ThrottleMs    = 300,
  [int]    $MaxRetries    = 10,
  [string] $SpaceKeysFile = ""
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ============================================================
# INITIALISATION
# ============================================================
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts      = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("HabilitationsConfluence_{0}.log" -f $ts)
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
  [System.IO.File]::AppendAllText($script:outCsv, "$line`r`n", [System.Text.Encoding]::UTF8)
}

function Format-DateFr([string]$isoDate) {
  if (-not $isoDate) { return "" }
  try {
    $dt = [DateTimeOffset]::Parse($isoDate)
    return $dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
  } catch { return $isoDate }
}

# ============================================================
# RESEAU — Proxy entreprise + TLS
# ============================================================
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# ============================================================
# HTTP HELPER avec decodage UTF-8 et retry
# ============================================================
function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{
        Method         = $Method
        Uri            = $Url
        Headers        = $Headers
        UseBasicParsing = $true
        ErrorAction    = "Stop"
      }
      if ($Body) {
        $params["ContentType"] = "application/json; charset=utf-8"
        $params["Body"]        = [System.Text.Encoding]::UTF8.GetBytes($Body)
      }

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

      if ($attempt -gt $MaxRetries) {
        return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
      }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6, $attempt)))
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

# --- Org Admin ---
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }
$orgData   = Import-Clixml -Path $orgCredFile
$orgId     = [string]$orgData.OrgId
$orgApiKey = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password
$orgHeaders = @{ Authorization = "Bearer $orgApiKey"; Accept = "application/json" }

# --- Site Admin ---
$siteCredFile = Join-Path $SecretsDir "site-admin.xml"
if (-not (Test-Path $siteCredFile)) {
  Log "Fichier site-admin.xml introuvable, creation interactive..." "WARN"
  [System.Windows.Forms.MessageBox]::Show(
    ("Le fichier site-admin.xml n'existe pas.`n`nVous allez devoir fournir :`n" +
     "  1. L'URL de votre site (ex: monsite.atlassian.net)`n" +
     "  2. Votre email administrateur`n  3. Un API Token"),
    "Creation credentials", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  $inputUrl = Read-Host "URL du site Atlassian (ex: monsite.atlassian.net)"
  $inputUrl = $inputUrl -replace "^https?://", "" -replace "/wiki.*$", "" -replace "/$", ""
  $adminEmail    = Read-Host "Email administrateur"
  $apiTokenSecure = Read-Host "API Token" -AsSecureString
  @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $siteCredFile
  Log "Credentials sauvegardes dans $siteCredFile"
}

$siteData  = Import-Clixml -Path $siteCredFile
$siteUrl   = [string]$siteData.SiteUrl
$siteEmail = [string]$siteData.Email
$siteToken = [System.Net.NetworkCredential]::new("", $siteData.ApiTokenSecureString).Password
$base64Auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${siteEmail}:${siteToken}"))
$confluenceHeaders = @{ Authorization = "Basic $base64Auth"; Accept = "application/json" }
$baseUrl = "https://$siteUrl"

Log "Site Atlassian: $baseUrl"

# ============================================================
# CACHES
# ============================================================
$script:groupMembersCache = @{}
$script:userEmailCache    = @{}
$script:userTypeCache     = @{}

# ============================================================
# FONCTIONS METIER
# ============================================================

function Get-TypeUserFromSubject($userObj) {
  $accType = ""; if ($userObj.accountType) { $accType = [string]$userObj.accountType }
  $isGuest = $false; $isExternal = $false
  if ($userObj.isGuest -eq $true) { $isGuest = $true }
  if ($userObj.isExternalCollaborator -eq $true) { $isExternal = $true }
  if ($accType -eq "app") { return "App" }
  if ($isGuest) { return "Guest" }
  if ($isExternal) { return "Externe" }
  if ($accType -eq "atlassian") { return "Interne" }
  return "Inconnu"
}

# --- Email Source 2 : Org Admin ---
function Get-UserEmailFromOrg([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/profile"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:orgHeaders
  $email = ""
  if ($resp.ok) {
    try {
      $json = $resp.content | ConvertFrom-Json
      if ($json.account -and $json.account.email) { $email = [string]$json.account.email }
      elseif ($json.email) { $email = [string]$json.email }
    } catch {}
  }
  return $email
}

# --- Email Source 3 : Confluence REST ---
function Get-ConfluenceUserInfo([string]$AccountId) {
  $url = "${script:baseUrl}/wiki/rest/api/user?accountId=${AccountId}"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:confluenceHeaders
  $info = @{ email = ""; typeUser = "Inconnu" }
  if ($resp.ok) {
    try {
      $json = $resp.content | ConvertFrom-Json
      if ($json.email) { $info.email = [string]$json.email }
      $info.typeUser = Get-TypeUserFromSubject $json
    } catch {}
  }
  return $info
}

# --- Email Source 4 : Jira REST ---
function Get-UserEmailFromJira([string]$AccountId) {
  $url = "${script:baseUrl}/rest/api/3/user?accountId=${AccountId}"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:confluenceHeaders
  $email = ""
  if ($resp.ok) {
    try {
      $json = $resp.content | ConvertFrom-Json
      if ($json.emailAddress) { $email = [string]$json.emailAddress }
    } catch {}
  }
  return $email
}

# --- Email cascade 4 sources ---
function Resolve-UserEmail([string]$AccountId, [string]$KnownEmail) {
  if ($KnownEmail) {
    if (-not $script:userEmailCache.ContainsKey($AccountId)) { $script:userEmailCache[$AccountId] = $KnownEmail }
    return $KnownEmail
  }
  if ($script:userEmailCache.ContainsKey($AccountId)) { return $script:userEmailCache[$AccountId] }

  $email = Get-UserEmailFromOrg -AccountId $AccountId
  if ($email) { $script:userEmailCache[$AccountId] = $email; return $email }

  $confInfo = Get-ConfluenceUserInfo -AccountId $AccountId
  if ($confInfo.typeUser -ne "Inconnu" -and -not $script:userTypeCache.ContainsKey($AccountId)) {
    $script:userTypeCache[$AccountId] = $confInfo.typeUser
  }
  if ($confInfo.email) { $script:userEmailCache[$AccountId] = $confInfo.email; return $confInfo.email }

  $email = Get-UserEmailFromJira -AccountId $AccountId
  if ($email) {
    $script:userEmailCache[$AccountId] = $email
    Log ("    Email trouve via Jira API pour {0}" -f $AccountId)
    return $email
  }

  $script:userEmailCache[$AccountId] = ""
  Log ("    Email introuvable pour {0} (4 sources epuisees)" -f $AccountId) "WARN"
  Start-Sleep -Milliseconds 100
  return ""
}

# --- TypeUser resolution ---
function Resolve-TypeUser([string]$AccountId, [string]$KnownType) {
  if ($KnownType -and $KnownType -ne "Inconnu") {
    if (-not $script:userTypeCache.ContainsKey($AccountId)) { $script:userTypeCache[$AccountId] = $KnownType }
    return $KnownType
  }
  if ($script:userTypeCache.ContainsKey($AccountId)) { return $script:userTypeCache[$AccountId] }

  $confInfo = Get-ConfluenceUserInfo -AccountId $AccountId
  if ($confInfo.email -and -not $script:userEmailCache.ContainsKey($AccountId)) {
    $script:userEmailCache[$AccountId] = $confInfo.email
  }
  $script:userTypeCache[$AccountId] = $confInfo.typeUser
  return $confInfo.typeUser
}

# --- Membres d'un groupe ---
function Get-GroupMembers([string]$GroupName) {
  if ($script:groupMembersCache.ContainsKey($GroupName)) { return $script:groupMembersCache[$GroupName] }

  $members = New-Object System.Collections.Generic.List[object]
  $encodedName = [System.Uri]::EscapeDataString($GroupName)
  $startAt = 0; $pageSize = 200

  while ($true) {
    $url = "${script:baseUrl}/wiki/rest/api/group/${encodedName}/member?limit=${pageSize}&start=${startAt}&expand=status"
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:confluenceHeaders
    if (-not $resp.ok) {
      Log ("    Erreur membres groupe '{0}' : status={1}" -f $GroupName, $resp.status) "WARN"
      break
    }
    $json = $resp.content | ConvertFrom-Json
    $countInPage = 0
    foreach ($member in $json.results) {
      $accId = [string]$member.accountId
      $dName = [string]$member.displayName
      $mEmail = ""; if ($member.email) { $mEmail = [string]$member.email }
      if ($mEmail -and -not $script:userEmailCache.ContainsKey($accId)) {
        $script:userEmailCache[$accId] = $mEmail
      }
      $mType = "Inconnu"
      if ($null -ne $member.accountType -or $null -ne $member.isGuest -or $null -ne $member.isExternalCollaborator) {
        $mType = Get-TypeUserFromSubject $member
      }
      if ($mType -ne "Inconnu" -and -not $script:userTypeCache.ContainsKey($accId)) {
        $script:userTypeCache[$accId] = $mType
      }
      $members.Add(@{ accountId=$accId; displayName=$dName; email=$mEmail; typeUser=$mType }) | Out-Null
      $countInPage++
    }
    if ($countInPage -lt $pageSize) { break }
    $startAt += $pageSize; Start-Sleep -Milliseconds 100
  }

  $script:groupMembersCache[$GroupName] = $members
  Log ("    Groupe '{0}' : {1} membres (cache)" -f $GroupName, $members.Count)
  return $members
}

# --- Niveau d'acces ---
function Get-NiveauAcces([string[]]$Operations) {
  $a=$false; $d=$false; $w=$false; $e=$false; $r=$false
  foreach ($op in $Operations) {
    if ($op -match "administer") { $a=$true }
    elseif ($op -match "delete|remove|archive") { $d=$true }
    elseif ($op -match "create|update") { $w=$true }
    elseif ($op -match "export") { $e=$true }
    elseif ($op -match "read|use") { $r=$true }
  }
  if ($a) { return "Admin" }
  if ($d) { return "Ecriture+Suppression" }
  if ($w) { return "Ecriture" }
  if ($e) { return "Lecture+Export" }
  if ($r) { return "Lecture" }
  return "Autre"
}

# --- Derniere modification via CQL ---
function Get-SpaceLastModified([string]$SpaceKey) {
  $cql = 'space = "' + $SpaceKey + '" order by lastModified desc'
  $encoded = [System.Uri]::EscapeDataString($cql)
  $url = "${script:baseUrl}/wiki/rest/api/content/search?cql=${encoded}&limit=1&expand=version"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:confluenceHeaders
  if ($resp.ok) {
    try {
      $json = $resp.content | ConvertFrom-Json
      if ($json.results -and $json.results.Count -gt 0 -and
          $json.results[0].version -and $json.results[0].version.when) {
        return Format-DateFr ([string]$json.results[0].version.when)
      }
    } catch {}
  }
  return ""
}

# --- Resoudre le nom d'un groupe a partir de son groupId ---
function Resolve-GroupName([string]$GroupId) {
  $url = "${script:baseUrl}/rest/api/3/group/bulk?groupId=${GroupId}"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $script:confluenceHeaders
  if ($resp.ok) {
    try {
      $json = $resp.content | ConvertFrom-Json
      if ($json.values -and $json.values.Count -gt 0) {
        return [string]$json.values[0].name
      }
    } catch {}
  }
  return $GroupId
}

# ==================== MAIN ====================
Log "=== Listing des habilitations Confluence ==="
Log "Site: $baseUrl"

# --- CSV header ---
$script:outCsv = Join-Path $ExportsDir ("HabilitationsConfluence_{0}.csv" -f $ts)
$allColumns = @(
  "SpaceKey","SpaceName","SpaceType","DateCreation","DateDerniereModif",
  "Groupe","TypeAcces","AccountId","Email","DisplayName",
  "TypeUser","NiveauAcces","OperationsDetail"
)
$headerLine = ($allColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($script:outCsv, "$headerLine`r`n", $utf8Bom)
Log "Export -> $($script:outCsv)"

# ==================== ETAPE 1 : Lister les espaces ====================
Log "=== ETAPE 1 : Recuperation des espaces Confluence ==="
$allSpaces = New-Object System.Collections.Generic.List[object]

if ($SpaceKeysFile -and (Test-Path $SpaceKeysFile)) {
  # ---- Mode fichier : uniquement les espaces listes ----
  Log "Mode fichier : $SpaceKeysFile"
  $fileKeys = Get-Content -Path $SpaceKeysFile -Encoding UTF8 | ForEach-Object {
    $line = $_.Trim()
    if ($line -and -not $line.StartsWith("#")) { ($line -split '\s+')[0] }
  } | Where-Object { $_ }

  foreach ($sk in $fileKeys) {
    $url = "${baseUrl}/wiki/api/v2/spaces?keys=$([System.Uri]::EscapeDataString($sk))"
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $confluenceHeaders
    if ($resp.ok) {
      $json = $resp.content | ConvertFrom-Json
      if ($json.results -and $json.results.Count -gt 0) {
        $sp = $json.results[0]
        $allSpaces.Add(@{
          key  = [string]$sp.key
          name = [string]$sp.name
          type = [string]$sp.type
          id   = [string]$sp.id
        }) | Out-Null
        Log ("  {0} => {1} (ID: {2})" -f $sp.key, $sp.name, $sp.id)
      } else {
        Log ("  {0} => NON TROUVE sur l'instance" -f $sk) "WARN"
      }
    } else {
      Log ("  {0} => ERREUR status={1}" -f $sk, $resp.status) "ERROR"
    }
    Start-Sleep -Milliseconds 150
  }
}
else {
  # ---- Mode complet : espaces globaux uniquement ----
  Log "Mode complet : espaces globaux uniquement (exclut les personnels)"
  $startAt = 0; $pageSize = 25
  while ($true) {
    $url = "${baseUrl}/wiki/rest/api/space?type=global&limit=${pageSize}&start=${startAt}"
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $confluenceHeaders
    if (-not $resp.ok) {
      Log ("Erreur API space start={0} : status={1}" -f $startAt, $resp.status) "ERROR"
      break
    }
    $json = $resp.content | ConvertFrom-Json
    $countInPage = 0
    foreach ($sp in $json.results) {
      $allSpaces.Add(@{
        key  = [string]$sp.key
        name = [string]$sp.name
        type = [string]$sp.type
        id   = ""
      }) | Out-Null
      $countInPage++
    }
    Log ("  Espaces charges : {0} (page start={1})" -f $allSpaces.Count, $startAt)
    if ($countInPage -lt $pageSize) { break }
    $startAt += $pageSize; Start-Sleep -Milliseconds 200
  }
}

$globalCount   = ($allSpaces | Where-Object { $_.type -eq "global" }).Count
$personalCount = ($allSpaces | Where-Object { $_.type -eq "personal" }).Count
Log ("Total espaces a auditer : {0} (global: {1}, personal: {2})" -f $allSpaces.Count, $globalCount, $personalCount)

if ($allSpaces.Count -eq 0) {
  Log "Aucun espace trouve." "ERROR"
  [System.Windows.Forms.MessageBox]::Show("Aucun espace trouve.", "Erreur", 0, 16) | Out-Null
  return
}

# ==================== ETAPE 2 : Permissions par espace (API v2) ====================
Log "=== ETAPE 2 : Recuperation des permissions par espace (API v2) ==="
$cSpaces = 0; $cLignesCsv = 0; $cErrPerm = 0; $cGone = 0
$cGuests = 0; $cExternes = 0; $cApps = 0; $cEmailManquant = 0
$startTime = Get-Date

# Cache pour la resolution groupId -> groupName
$script:groupNameCache = @{}

foreach ($space in $allSpaces) {
  $cSpaces++
  $spKey  = $space.key
  $spName = $space.name
  $spType = $space.type

  # --- Progression ---
  $elapsed = (Get-Date) - $startTime
  if ($cSpaces -gt 1) {
    $remaining = [Math]::Round(($elapsed.TotalSeconds / ($cSpaces - 1)) * ($allSpaces.Count - $cSpaces) / 60, 1)
    $etaMsg = "~${remaining} min restantes"
  } else { $etaMsg = "calcul..." }
  Write-Progress -Activity "Permissions Confluence - $etaMsg" `
    -Status ("{0}/{1} {2}" -f $cSpaces, $allSpaces.Count, $spKey) `
    -PercentComplete ([int](100 * $cSpaces / $allSpaces.Count))

  # --- Resoudre le Space ID si pas deja connu ---
  $spaceId = $space.id
  if (-not $spaceId) {
    $encodedKey = [System.Uri]::EscapeDataString($spKey)
    $urlResolve = "${baseUrl}/wiki/api/v2/spaces?keys=${encodedKey}"
    $respResolve = Invoke-ApiCall -Method "GET" -Url $urlResolve -Headers $confluenceHeaders
    if ($respResolve.ok) {
      $jsonResolve = $respResolve.content | ConvertFrom-Json
      if ($jsonResolve.results -and $jsonResolve.results.Count -gt 0) {
        $spaceId = [string]$jsonResolve.results[0].id
      }
    }
  }

  if (-not $spaceId) {
    Log ("{0}/{1} {2} ({3}) : impossible de resoudre l'ID" -f $cSpaces, $allSpaces.Count, $spKey, $spName) "ERROR"
    $cErrPerm++; continue
  }

  # --- Date de creation via API v1 expand=history (fonctionne sans probleme) ---
  $dateCreation = ""
  $encodedKey = [System.Uri]::EscapeDataString($spKey)
  $histUrl = "${baseUrl}/wiki/rest/api/space/${encodedKey}?expand=history"
  $respHist = Invoke-ApiCall -Method "GET" -Url $histUrl -Headers $confluenceHeaders
  if ($respHist.ok) {
    try {
      $jsonHist = $respHist.content | ConvertFrom-Json
      if ($jsonHist.history -and $jsonHist.history.createdDate) {
        $dateCreation = Format-DateFr ([string]$jsonHist.history.createdDate)
      }
    } catch {}
  }

  # --- Date derniere modification (CQL) ---
  $dateLastModif = Get-SpaceLastModified -SpaceKey $spKey

  # --- Permissions via API v2 (pas de 405) ---
  $allPermissions = New-Object System.Collections.Generic.List[object]
  $permUrl = "${baseUrl}/wiki/api/v2/spaces/${spaceId}/permissions"
  $permResp = Invoke-ApiCall -Method "GET" -Url $permUrl -Headers $confluenceHeaders

  if (-not $permResp.ok) {
    if ($permResp.status -eq 410) {
      Log ("{0}/{1} {2} ({3}) : espace supprime/archive (410)" -f $cSpaces, $allSpaces.Count, $spKey, $spName) "WARN"
      $cGone++
    } else {
      Log ("  ERREUR permissions {0} : status={1} {2}" -f $spKey, $permResp.status, $permResp.error) "ERROR"
    }
    $cErrPerm++; continue
  }

  # Collecter toutes les permissions (avec pagination)
  $jsonPerms = $permResp.content | ConvertFrom-Json
  if ($jsonPerms.results) {
    foreach ($p in $jsonPerms.results) { $allPermissions.Add($p) | Out-Null }
  }
  # Pagination
  while ($jsonPerms._links -and $jsonPerms._links.next) {
    $nextUrl = $jsonPerms._links.next
    # Si l'URL est relative, prefixer
    if ($nextUrl -notmatch "^https?://") {
      $nextUrl = "${baseUrl}/wiki${nextUrl}"
    }
    $respNext = Invoke-ApiCall -Method "GET" -Url $nextUrl -Headers $confluenceHeaders
    if (-not $respNext.ok) { break }
    $jsonPerms = $respNext.content | ConvertFrom-Json
    if ($jsonPerms.results) {
      foreach ($p in $jsonPerms.results) { $allPermissions.Add($p) | Out-Null }
    }
  }

  if ($allPermissions.Count -eq 0) {
    Log ("{0}/{1} {2} : aucune permission" -f $cSpaces, $allSpaces.Count, $spKey)
    continue
  }

  # --- Regrouper les operations par sujet (format API v2) ---
  $subjects = @{}

  foreach ($perm in $allPermissions) {
    # Operation
    $opKey = ""
    if ($perm.operation) {
      $opKey = "{0}/{1}" -f $perm.operation.key, $perm.operation.target
    }

    $principal = $perm.principal
    if (-not $principal) {
      # Fallback : format alternatif API v2
      if ($perm.subject) { $principal = $perm.subject }
      else { continue }
    }

    $principalType = [string]$principal.type
    $principalId   = [string]$principal.id

    if ($principalType -eq "group") {
      # Resoudre le nom du groupe (l'API v2 retourne souvent un groupId)
      $gName = $principalId
      if ($script:groupNameCache.ContainsKey($principalId)) {
        $gName = $script:groupNameCache[$principalId]
      } else {
        $resolvedName = Resolve-GroupName -GroupId $principalId
        $script:groupNameCache[$principalId] = $resolvedName
        $gName = $resolvedName
      }

      $sk = "group:$gName"
      if (-not $subjects.ContainsKey($sk)) {
        $subjects[$sk] = @{
          type       = "group"
          name       = $gName
          operations = New-Object System.Collections.Generic.List[string]
        }
      }
      if ($opKey -and -not $subjects[$sk].operations.Contains($opKey)) {
        $subjects[$sk].operations.Add($opKey)
      }
    }
    elseif ($principalType -eq "user") {
      $sk = "user:$principalId"
      if (-not $subjects.ContainsKey($sk)) {
        $subjects[$sk] = @{
          type        = "user"
          accountId   = $principalId
          displayName = ""
          email       = ""
          typeUser    = "Inconnu"
          operations  = New-Object System.Collections.Generic.List[string]
        }
      }
      if ($opKey -and -not $subjects[$sk].operations.Contains($opKey)) {
        $subjects[$sk].operations.Add($opKey)
      }
    }
    elseif ($principalType -eq "role") {
      # Roles systeme (ex: "anonymous", "known") — traiter comme un groupe special
      $roleName = if ($principal.name) { [string]$principal.name } else { $principalId }
      $sk = "role:$roleName"
      if (-not $subjects.ContainsKey($sk)) {
        $subjects[$sk] = @{
          type       = "role"
          name       = $roleName
          operations = New-Object System.Collections.Generic.List[string]
        }
      }
      if ($opKey -and -not $subjects[$sk].operations.Contains($opKey)) {
        $subjects[$sk].operations.Add($opKey)
      }
    }
  }

  $groupCount = ($subjects.Values | Where-Object { $_.type -eq "group" }).Count
  $userCount  = ($subjects.Values | Where-Object { $_.type -eq "user" }).Count
  $roleCount  = ($subjects.Values | Where-Object { $_.type -eq "role" }).Count
  $lignesEspace = 0

  foreach ($subjectKey in $subjects.Keys) {
    $subject = $subjects[$subjectKey]
    $ops       = $subject.operations | Sort-Object
    $opsDetail = ($ops -join ", ")
    $niveau    = Get-NiveauAcces -Operations $ops

    if ($subject.type -eq "group") {
      $members = Get-GroupMembers -GroupName $subject.name

      if ($members.Count -eq 0) {
        $out = [ordered]@{
          SpaceKey=$spKey; SpaceName=$spName; SpaceType=$spType
          DateCreation=$dateCreation; DateDerniereModif=$dateLastModif
          Groupe=$subject.name; TypeAcces="Groupe"
          AccountId=""; Email=""; DisplayName="(groupe vide)"
          TypeUser=""; NiveauAcces=$niveau; OperationsDetail=$opsDetail
        }
        $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
        Write-CsvLine $csvLine
        $cLignesCsv++; $lignesEspace++
      } else {
        foreach ($member in $members) {
          $email = Resolve-UserEmail -AccountId $member.accountId -KnownEmail $member.email
          if (-not $email) { $cEmailManquant++ }
          $typeUser = Resolve-TypeUser -AccountId $member.accountId -KnownType $member.typeUser
          switch ($typeUser) { "Guest" { $cGuests++ } "Externe" { $cExternes++ } "App" { $cApps++ } }

          $out = [ordered]@{
            SpaceKey=$spKey; SpaceName=$spName; SpaceType=$spType
            DateCreation=$dateCreation; DateDerniereModif=$dateLastModif
            Groupe=$subject.name; TypeAcces="Groupe"
            AccountId=$member.accountId; Email=$email; DisplayName=$member.displayName
            TypeUser=$typeUser; NiveauAcces=$niveau; OperationsDetail=$opsDetail
          }
          $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
          Write-CsvLine $csvLine
          $cLignesCsv++; $lignesEspace++
        }
      }
    }
    elseif ($subject.type -eq "user") {
      $email = Resolve-UserEmail -AccountId $subject.accountId -KnownEmail $subject.email
      if (-not $email) { $cEmailManquant++ }
      $typeUser = Resolve-TypeUser -AccountId $subject.accountId -KnownType $subject.typeUser
      switch ($typeUser) { "Guest" { $cGuests++ } "Externe" { $cExternes++ } "App" { $cApps++ } }

      $out = [ordered]@{
        SpaceKey=$spKey; SpaceName=$spName; SpaceType=$spType
        DateCreation=$dateCreation; DateDerniereModif=$dateLastModif
        Groupe=""; TypeAcces="Nominatif"
        AccountId=$subject.accountId; Email=$email; DisplayName=$subject.displayName
        TypeUser=$typeUser; NiveauAcces=$niveau; OperationsDetail=$opsDetail
      }
      $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
      Write-CsvLine $csvLine
      $cLignesCsv++; $lignesEspace++
    }
    elseif ($subject.type -eq "role") {
      $out = [ordered]@{
        SpaceKey=$spKey; SpaceName=$spName; SpaceType=$spType
        DateCreation=$dateCreation; DateDerniereModif=$dateLastModif
        Groupe=$subject.name; TypeAcces="Role"
        AccountId=""; Email=""; DisplayName=$subject.name
        TypeUser="Systeme"; NiveauAcces=$niveau; OperationsDetail=$opsDetail
      }
      $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
      Write-CsvLine $csvLine
      $cLignesCsv++; $lignesEspace++
    }
  }

  Log ("{0}/{1} {2} ({3}) : cree={4} modif={5} | {6} grp, {7} nom, {8} role => {9} lignes" -f `
    $cSpaces, $allSpaces.Count, $spKey, $spName, $dateCreation, $dateLastModif, `
    $groupCount, $userCount, $roleCount, $lignesEspace)

  Start-Sleep -Milliseconds $ThrottleMs
}

Write-Progress -Activity "Permissions Confluence" -Completed

# ============================================================
# RESUME
# ============================================================
Log "============================================"
Log "RESUME"
Log "============================================"
Log "Site: $baseUrl"
Log ("Espaces analyses : {0} (global: {1}, personal: {2})" -f $cSpaces, $globalCount, $personalCount)
Log ("Espaces supprimes/archives (410) : {0}" -f $cGone)
Log ("Autres erreurs permissions : {0}" -f ($cErrPerm - $cGone))
Log ("Groupes uniques en cache : {0}" -f $script:groupMembersCache.Count)
Log ("Utilisateurs uniques en cache : {0}" -f $script:userEmailCache.Count)
Log ("Lignes CSV generees : {0}" -f $cLignesCsv)
Log "---"
Log ("Guests (lignes) : {0}" -f $cGuests)
Log ("Externes (lignes) : {0}" -f $cExternes)
Log ("Apps (lignes) : {0}" -f $cApps)
Log ("Emails manquants (lignes) : {0}" -f $cEmailManquant)
$totalMin = [Math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
Log "Duree totale : $totalMin minutes"
Log "============================================"
Log "Termine."

# --- Popup resume ---
$summaryMsg = @"
Listing des habilitations Confluence termine.

Site: $baseUrl
Espaces : $cSpaces (global: $globalCount, personal: $personalCount)
Espaces 410 : $cGone
Autres erreurs : $($cErrPerm - $cGone)
Groupes uniques : $($script:groupMembersCache.Count)
Utilisateurs uniques : $($script:userEmailCache.Count)
Lignes CSV : $cLignesCsv

Guests : $cGuests
Externes : $cExternes
Apps : $cApps
Emails manquants : $cEmailManquant

Duree : $totalMin min
Export : $($script:outCsv)
"@

[System.Windows.Forms.MessageBox]::Show(
  $summaryMsg,
  "Resultat - Habilitations Confluence",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
