#Requires -Version 5.1
<#
.SYNOPSIS
  Desactivation-NoAccess.ps1
  Desactive les comptes actifs sans acces applicatif, crees depuis plus de 60 jours.

.DESCRIPTION
  Ce script :
  1. Liste les users via GET /admin/v1/orgs/{orgId}/users (pagine)
  2. Filtre par domaines email autorises
  3. Identifie les comptes actifs avec "No app access" (product_access vide)
  4. Verifie la date de creation via l'API Events (event user_created)
  5. En DRYRUN : enrichit avec le statut CMDB Assets (Statut, Date Sortie, Motif)
  6. Desactive les comptes crees depuis plus de 60 jours (en mode EXECUTE)

  Le CSV est ecrit AU FIL DE L'EAU : consultable des le debut.

  DOMAINES AUTORISES :
    harmonie-mutuelle.fr, prestataire.sihm.fr, mutex.fr,
    mutex-exterieur.fr, chorum.fr

  ENRICHISSEMENT CMDB (mode DRYRUN) :
    - Endpoint : /gateway/api/jsm/assets/workspace/{wId}/v1/object/aql
    - Types : "Employe", "Prestataire" (schema Referentiel personne, key=RP)
    - Recherche par "Compte Jira" (accountId)
    - Attributs par ID : 408=Statut, 380=Date Sortie, 381=Motif Sortie, 379=Date Entree
    - Permet de valider si la personne est sortie (Inactif) ou toujours en poste (Actif)

  MODE DRYRUN  : liste les comptes eligibles + statut CMDB sans les desactiver.
  MODE EXECUTE : desactive reellement les comptes eligibles.

  COLONNES DU CSV DE SORTIE :
    AccountId      => ID du compte Atlassian
    Email          => Email du compte
    Name           => Nom affiche
    Domain         => Domaine email
    AccountStatus  => Statut du compte
    AccessBillable => True/False
    NbProducts     => Nombre de produits accessibles
    DateCreation   => Date de creation du compte (DD/MM/YYYY HH:mm)
    JoursDepuis    => Nombre de jours depuis la creation
    Eligible       => Oui/Non
    StatutCMDB     => Statut dans la CMDB Assets (Actif/Inactif/Non trouve)
    DateSortieCMDB => Date de sortie selon la CMDB
    MotifSortie    => Motif de sortie (CMDB)
    ResultAction   => Resultat de l'action

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 1.2 -- Aout 2026
  Compatibilite  : PowerShell 5.1+
  Prerequis      : secrets\org-admin.xml (OrgId + API Key)
                   secrets\site-admin.xml (Jiradot -- pour acces Assets)

  APIs UTILISEES :
    GET  /admin/v1/orgs/{orgId}/users               (liste paginee)
    GET  /admin/v1/orgs/{orgId}/events?q={email}    (date de creation via user_created)
    POST /users/{accountId}/manage/lifecycle/disable (desactivation)
    POST /gateway/api/jsm/assets/workspace/{wId}/v1/object/aql (CMDB)

  CRITERES D'ELIGIBILITE :
    1. Domaine email autorise
    2. Compte actif (account_status = "active")
    3. Aucun acces applicatif (product_access vide)
    4. Cree depuis plus de 60 jours (parametrable via -SeuilJours)

.EXAMPLE
  .\Desactivation-NoAccess.ps1
  .\Desactivation-NoAccess.ps1 -SeuilJours 90
#>

[CmdletBinding()]
param(
  [int] $ThrottleMs  = 400,
  [int] $MaxRetries  = 10,
  [int] $SeuilJours  = 60
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# -------------------- Domaines autorises --------------------
$domainesAutorises = @(
  "harmonie-mutuelle.fr",
  "prestataire.sihm.fr",
  "mutex.fr",
  "mutex-exterieur.fr",
  "chorum.fr"
)

# -------------------- Date de reference --------------------
$dateRef    = (Get-Date).Date
$dateSeuil  = $dateRef.AddDays(-$SeuilJours)

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts      = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("DesactivationNoAccess_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level = "INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# -------------------- Parsing de dates robuste --------------------
function ConvertTo-DateSafe([string]$dateStr) {
  if ([string]::IsNullOrWhiteSpace($dateStr)) { return [datetime]::MinValue }

  try { return [datetime]::Parse($dateStr, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}

  $formats = @(
    "MM/dd/yyyy HH:mm:ss",
    "MM/dd/yyyy",
    "dd/MM/yyyy HH:mm:ss",
    "dd/MM/yyyy",
    "yyyy-MM-dd'T'HH:mm:ss.fff'Z'",
    "yyyy-MM-dd'T'HH:mm:ss'Z'",
    "yyyy-MM-dd'T'HH:mm:sszzz",
    "yyyy-MM-dd",
    "dd/MMM/yy",
    "dd/MMM/yy h:mm tt",
    "dd/MMM/yy H:mm",
    "dd/MMM/yyyy",
    "dd/MMM/yyyy h:mm tt"
  )
  foreach ($fmt in $formats) {
    try {
      return [datetime]::ParseExact($dateStr, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {}
  }
  return [datetime]::MinValue
}

function ConvertTo-DateFromAsset([string]$assetDate) {
  if ([string]::IsNullOrWhiteSpace($assetDate)) { return "" }
  $dt = ConvertTo-DateSafe $assetDate
  if ($dt -eq [datetime]::MinValue) { return $assetDate }
  return $dt.ToString("yyyy-MM-dd")
}

# -------------------- Popup : choix DryRun / Execute / Annuler --------------------
function Choose-Mode {
  $result = [System.Windows.Forms.MessageBox]::Show(
    ("Desactivation des comptes 'No app access' crees depuis plus de $SeuilJours jours.`n" +
     "Domaines : $($domainesAutorises -join ', ')`n`n" +
     "Oui     = EXECUTE (desactivations REELLES)`n" +
     "Non     = DRY RUN (liste les comptes eligibles + statut CMDB)`n" +
     "Annuler = Quitter"),
    "Mode d'execution - Desactivation No Access",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return "EXECUTE" }
  if ($result -eq [System.Windows.Forms.DialogResult]::No)  { return "DRYRUN"  }
  return "CANCEL"
}

# -------------------- Load credentials (org-admin) --------------------
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }

$data        = Import-Clixml -Path $orgCredFile
$orgId       = [string]$data.OrgId
$apiKeyPlain = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password
$headers     = @{
  Authorization  = "Bearer $apiKeyPlain"
  Accept         = "application/json"
  "Content-Type" = "application/json"
}

# -------------------- Site credentials (pour Assets CMDB) --------------------
$siteCredFile = Join-Path $SecretsDir "site-admin.xml"
$siteHeaders  = $null
$siteBaseUrl  = ""

if (Test-Path $siteCredFile) {
  $siteData  = Import-Clixml -Path $siteCredFile
  $siteUrl   = [string]$siteData.SiteUrl
  $siteEmail = [string]$siteData.Email
  $siteToken = [System.Net.NetworkCredential]::new("", $siteData.ApiTokenSecureString).Password
  $siteAuth  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("${siteEmail}:${siteToken}"))
  $siteHeaders = @{
    Authorization = "Basic $siteAuth"
    Accept        = "application/json"
  }
  $siteBaseUrl = "https://$siteUrl"
} else {
  Log "site-admin.xml non trouve -- enrichissement CMDB desactive" "WARN"
}

# -------------------- Network --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# -------------------- HTTP helper --------------------
function Invoke-ApiCall {
  param(
    [string]    $Method,
    [string]    $Url,
    [hashtable] $CustomHeaders = $null,
    [string]    $Body          = $null
  )

  $hdrs    = if ($CustomHeaders) { $CustomHeaders } else { $script:headers }
  $attempt = 0

  while ($true) {
    $attempt++
    try {
      $params = @{
        Method          = $Method
        Uri             = $Url
        Headers         = $hdrs
        UseBasicParsing = $true
        ErrorAction     = "Stop"
      }
      if ($Body) {
        $params["ContentType"] = "application/json; charset=utf-8"
        $params["Body"]        = [System.Text.Encoding]::UTF8.GetBytes($Body)
      }
      $resp = Invoke-WebRequest @params
      return @{ ok = $true; status = [int]$resp.StatusCode; content = $resp.Content }
    } catch {
      $status  = 0
      $errBody = ""
      try {
        $status  = [int]$_.Exception.Response.StatusCode
        $reader  = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $reader.ReadToEnd()
        $reader.Close()
      } catch {}

      if ($attempt -gt $MaxRetries) {
        return @{ ok = $false; status = $status; error = $_.Exception.Message; body = $errBody }
      }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6, $attempt)))
        Log "Retry $Method status=$status in ${sleepSec}s ($attempt/$MaxRetries)" "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }
      return @{ ok = $false; status = $status; error = $_.Exception.Message; body = $errBody }
    }
  }
}

# ============================================================
# JIRA ASSETS -- CMDB ENRICHMENT
# ============================================================

$assetsWorkspaceId = $null

function Get-AssetsWorkspaceId {
  if (-not $script:siteHeaders -or -not $script:siteBaseUrl) { return $null }
  $url  = "{0}/rest/servicedeskapi/assets/workspace" -f $script:siteBaseUrl
  $resp = Invoke-ApiCall -Method "GET" -Url $url -CustomHeaders $script:siteHeaders
  if ($resp.ok) {
    $json = $resp.content | ConvertFrom-Json
    if ($json.values -and $json.values.Count -gt 0) {
      return [string]$json.values[0].workspaceId
    }
  }
  return $null
}

function Search-AssetByAccountId([string]$AccountId) {
  if ([string]::IsNullOrWhiteSpace($AccountId) -or -not $script:assetsWorkspaceId) { return $null }
  if (-not $script:siteHeaders -or -not $script:siteBaseUrl) { return $null }

  $url = "{0}/gateway/api/jsm/assets/workspace/{1}/v1/object/aql" -f $script:siteBaseUrl, $script:assetsWorkspaceId
  $aql = 'objectType IN ("Employe", "Prestataire") AND "Compte Jira" = "{0}"' -f $AccountId

  $body = @{
    qlQuery           = $aql
    maxResults        = 1
    includeAttributes = $true
  } | ConvertTo-Json -Depth 5 -Compress

  $resp = Invoke-ApiCall -Method "POST" -Url $url -CustomHeaders $script:siteHeaders -Body $body
  if (-not $resp.ok) { return $null }

  $json    = $resp.content | ConvertFrom-Json
  $objects = $null
  if ($json.values)        { $objects = $json.values }
  elseif ($json.objectEntries) { $objects = $json.objectEntries }

  if (-not $objects -or ($objects | Measure-Object).Count -eq 0) { return $null }

  $asset  = $objects[0]
  $result = @{ Statut = ""; DateSortie = ""; MotifSortie = ""; DateEntree = "" }

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
        if     ($val.displayValue)              { $attrVal = [string]$val.displayValue }
        elseif ($val.value)                     { $attrVal = [string]$val.value }
        elseif ($val.status -and $val.status.name) { $attrVal = [string]$val.status.name }
      }

      switch ($attrId) {
        "408" { $result.Statut      = $attrVal }
        "380" { $result.DateSortie  = $attrVal }
        "381" { $result.MotifSortie = $attrVal }
        "379" { $result.DateEntree  = $attrVal }
      }
    }
  }
  return $result
}

# -------------------- API : Date de creation via Events --------------------
function Get-CreationDate([string]$Email, [string]$AccountId) {
  $encodedEmail = [System.Uri]::EscapeDataString($Email)
  $url  = "https://api.atlassian.com/admin/v1/orgs/$script:orgId/events?q=$encodedEmail&limit=50"
  $resp = Invoke-ApiCall -Method "GET" -Url $url

  if (-not $resp.ok) {
    Log "  Events API erreur pour $Email : status=$($resp.status)" "WARN"
    return $null
  }

  try {
    $json = $resp.content | ConvertFrom-Json
    if (-not $json.data) { return $null }

    foreach ($evt in $json.data) {
      $evtAction = [string]$evt.attributes.action
      if ($evtAction -eq "user_created") {
        $matchUser = $false
        foreach ($ctx in $evt.attributes.context) {
          if ($ctx.type -eq "users" -and $ctx.id -eq $AccountId) { $matchUser = $true; break }
        }
        if (-not $matchUser) {
          foreach ($ctn in $evt.attributes.container) {
            if ($ctn.type -eq "users" -and $ctn.id -eq $AccountId) { $matchUser = $true; break }
          }
        }
        if (-not $matchUser -and $evt.attributes.context.Count -eq 0) { $matchUser = $true }

        if ($matchUser) {
          $eventTime = ConvertTo-DateSafe $evt.attributes.time
          if ($eventTime -ne [datetime]::MinValue) { return $eventTime }
        }
      }
    }

    # Fallback : plus ancien event
    $oldest = $json.data | Sort-Object { (ConvertTo-DateSafe $_.attributes.time).Ticks } | Select-Object -First 1
    if ($oldest) {
      $oldestTime = ConvertTo-DateSafe $oldest.attributes.time
      if ($oldestTime -ne [datetime]::MinValue) {
        Log "  Pas de user_created pour $Email, utilisation du plus ancien event ($($oldest.attributes.action) du $($oldestTime.ToLocalTime().ToString('dd/MM/yyyy')))"
        return $oldestTime
      }
    }
  } catch {
    Log "  Erreur parsing events $Email : $($_.Exception.Message)" "WARN"
  }
  return $null
}

# -------------------- API : Disable --------------------
function Disable-User([string]$AccountId) {
  $url  = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/disable"
  $body = '{"message":"Desactivation automatique - compte sans acces applicatif depuis plus de 60 jours"}'
  return Invoke-ApiCall -Method "POST" -Url $url -Body $body
}

# -------------------- Extraction du domaine email --------------------
function Get-EmailDomain([string]$email) {
  if ($email -match "@(.+)$") { return $matches[1].ToLower() }
  return ""
}

# ==================== MAIN ====================

$mode = Choose-Mode
if ($mode -eq "CANCEL") { Log "Annule." "WARN"; return }

# Double confirmation en mode EXECUTE
if ($mode -eq "EXECUTE") {
  $confirm = [System.Windows.Forms.MessageBox]::Show(
    ("ATTENTION : Vous allez DESACTIVER les comptes`n" +
     "sans acces applicatif crees depuis plus de $SeuilJours jours.`n`n" +
     "Domaines concernes :`n" +
     ($domainesAutorises | ForEach-Object { "  - $_" } | Out-String) +
     "Confirmer la desactivation ?"),
    "CONFIRMATION DESACTIVATION",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Stop
  )
  if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
    Log "Desactivation annulee par l'utilisateur." "WARN"
    return
  }
}

Log "Mode=$mode | Seuil=$SeuilJours jours | Date seuil=$($dateSeuil.ToString('dd/MM/yyyy'))"
Log "Criteres: account_status=active AND product_access=vide AND cree avant $($dateSeuil.ToString('dd/MM/yyyy'))"
Log "Domaines autorises: $($domainesAutorises -join ', ')"
Log "Source date creation: Events API (event user_created)"

# -------------------- Decouverte Workspace Assets --------------------
if ($siteHeaders) {
  Log "Decouverte du workspace Assets..."
  $assetsWorkspaceId = Get-AssetsWorkspaceId
  if ($assetsWorkspaceId) {
    Log "  Workspace Assets : $assetsWorkspaceId"
    Log "  Enrichissement CMDB active (Statut, Date Sortie, Motif)"
  } else {
    Log "  Workspace Assets non trouve -- enrichissement CMDB desactive" "WARN"
  }
} else {
  Log "Pas de credentials site -- enrichissement CMDB desactive" "WARN"
}

# -------------------- Export incremental : header --------------------
$outCsv     = Join-Path $ExportsDir ("DesactivationNoAccess_Result_{0}_{1}.csv" -f $mode, $ts)
$allColumns = @("AccountId","Email","Name","Domain","AccountStatus","AccessBillable","NbProducts","DateCreation","JoursDepuis","Eligible","StatutCMDB","DateSortieCMDB","MotifSortie","ResultAction")
$headerLine = ($allColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
$headerLine | Set-Content -Path $outCsv -Encoding UTF8
Log "Export incremental -> $outCsv"

# Compteurs
$cTotalLus       = 0; $cDomainOk       = 0
$cActifAvecAcces = 0; $cNonActif       = 0; $cNoAccess    = 0
$cEligible       = 0; $cNonEligible    = 0; $cDateInconnue = 0
$cDisabled       = 0; $cError          = 0
$cAssetTrouve    = 0; $cAssetNonTrouve = 0

$pageNum  = 0
$nextUrl  = "https://api.atlassian.com/admin/v1/orgs/$orgId/users?limit=100"
$startTime = Get-Date

Log "=== Parcours des comptes de l'organisation ==="

while ($nextUrl) {
  $pageNum++

  $resp = Invoke-ApiCall -Method "GET" -Url $nextUrl
  if (-not $resp.ok) { throw "Erreur API users page $pageNum : status=$($resp.status) $($resp.error)" }

  $json         = $resp.content | ConvertFrom-Json
  $usersInPage  = $json.data
  $cTotalLus   += $usersInPage.Count
  $noAccessInPage = 0

  foreach ($user in $usersInPage) {
    $accId    = [string]$user.account_id
    $email    = [string]$user.email
    $name     = [string]$user.name
    $status   = [string]$user.account_status
    $billable = $user.access_billable
    $products = $user.product_access
    $nbProd   = 0
    if ($products) { $nbProd = @($products).Count }
    $domain   = Get-EmailDomain $email

    # Filtre domaine
    if ($domain -notin $domainesAutorises) { continue }
    $cDomainOk++

    $dateCreation   = ""
    $joursDepuis    = ""
    $eligible       = "Non"
    $resultAction   = ""
    $statutCMDB     = ""
    $dateSortieCMDB = ""
    $motifSortie    = ""

    # ---- NON ACTIF : ignorer ----
    if ($status -ne "active") {
      $cNonActif++
      $resultAction = "Compte non actif ($status)"
      $out = [ordered]@{
        AccountId=$accId; Email=$email; Name=$name; Domain=$domain
        AccountStatus=$status; AccessBillable=$billable; NbProducts=$nbProd
        DateCreation=""; JoursDepuis=""; Eligible="Non"
        StatutCMDB=""; DateSortieCMDB=""; MotifSortie=""
        ResultAction=$resultAction
      }
      $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
      Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8
      continue
    }

    # ---- ACTIF AVEC ACCES : ignorer ----
    if ($nbProd -gt 0) {
      $cActifAvecAcces++
      $resultAction = "Compte actif avec acces ($nbProd produits)"
      $out = [ordered]@{
        AccountId=$accId; Email=$email; Name=$name; Domain=$domain
        AccountStatus=$status; AccessBillable=$billable; NbProducts=$nbProd
        DateCreation=""; JoursDepuis=""; Eligible="Non"
        StatutCMDB=""; DateSortieCMDB=""; MotifSortie=""
        ResultAction=$resultAction
      }
      $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
      Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8
      continue
    }

    # ---- ACTIF SANS ACCES : chercher date de creation via Events API ----
    $cNoAccess++
    $noAccessInPage++

    Write-Progress -Activity "Verification ($mode)" -Status "Page $pageNum | $email" -PercentComplete -1

    $createdDate = Get-CreationDate -Email $email -AccountId $accId

    # ---- Enrichissement CMDB ----
    if ($assetsWorkspaceId) {
      $assetData = Search-AssetByAccountId -AccountId $accId
      if ($assetData) {
        $cAssetTrouve++
        $statutCMDB     = $assetData.Statut
        $dateSortieCMDB = ConvertTo-DateFromAsset $assetData.DateSortie
        $motifSortie    = $assetData.MotifSortie
      } else {
        $cAssetNonTrouve++
        $statutCMDB = "Non trouve"
      }
    }

    if ($createdDate) {
      $createdLocal = $createdDate.ToLocalTime()
      $dateCreation = $createdLocal.ToString("dd/MM/yyyy HH:mm")
      $jours        = [int]($dateRef - $createdLocal.Date).TotalDays
      $joursDepuis  = $jours

      if ($jours -ge $SeuilJours) {
        $eligible = "Oui"
        $cEligible++

        if ($mode -eq "EXECUTE") {
          Log "  Desactivation $email (cree le $dateCreation, $jours jours, 0 produit)..."
          $disResp = Disable-User -AccountId $accId
          if ($disResp.ok -or $disResp.status -eq 204) {
            $resultAction = "Desactive (cree depuis $jours jours, 0 produit)"
            $cDisabled++
            Log "  OK desactive $email"
          } else {
            $resultAction = "ERREUR DESACTIVATION: status=$($disResp.status)"
            $cError++
            Log "  ERREUR desactivation $email : status=$($disResp.status) body=$($disResp.body)" "ERROR"
          }
        } else {
          $cmdbInfo = ""
          if ($statutCMDB)     { $cmdbInfo  = " | CMDB=$statutCMDB" }
          if ($dateSortieCMDB) { $cmdbInfo += " Sortie=$dateSortieCMDB" }
          $resultAction = "[DRYRUN] A desactiver (cree le $dateCreation, $jours jours, 0 produit$cmdbInfo)"
          Log ("  ELIGIBLE: {0} cree le {1} ({2} jours), 0 produit | CMDB: Statut={3} Sortie={4} Motif={5}" -f $email, $dateCreation, $jours, $statutCMDB, $dateSortieCMDB, $motifSortie)
        }
      } else {
        $cNonEligible++
        $resultAction = "Non desactive (cree depuis $jours jours < $SeuilJours)"
        Log "  $email => cree depuis $jours jours < $SeuilJours, non eligible"
      }
    } else {
      $cDateInconnue++
      $resultAction = "Non desactive (date creation inconnue)"
      Log "  $email => No access, date creation inconnue"
    }

    # Ecriture incrementale
    $out = [ordered]@{
      AccountId=$accId; Email=$email; Name=$name; Domain=$domain
      AccountStatus=$status; AccessBillable=$billable; NbProducts=$nbProd
      DateCreation=$dateCreation; JoursDepuis=$joursDepuis
      Eligible=$eligible
      StatutCMDB=$statutCMDB; DateSortieCMDB=$dateSortieCMDB; MotifSortie=$motifSortie
      ResultAction=$resultAction
    }
    $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
    Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8

    Start-Sleep -Milliseconds $ThrottleMs
  }

  # Estimation temps
  $elapsed    = (Get-Date) - $startTime
  $elapsedMin = [Math]::Round($elapsed.TotalMinutes, 1)

  Log ("  Page {0}: {1} users, {2} no access | Total: {3} lus, {4} domaine OK, {5} actifs+acces, {6} non actifs, {7} no access ({8} eligibles) | {9} min" -f `
    $pageNum, $usersInPage.Count, $noAccessInPage, `
    $cTotalLus, $cDomainOk, $cActifAvecAcces, $cNonActif, $cNoAccess, $cEligible, $elapsedMin)

  # Pagination
  $nextUrl = $null
  if ($json.links -and $json.links.next) {
    $nextUrl = [string]$json.links.next
  }

  Start-Sleep -Milliseconds 200
}

# -------------------- Summary --------------------
Log "============================================"
Log "RESUME"
Log "============================================"
Log "Mode: $mode"
Log "Domaines: $($domainesAutorises -join ', ')"
Log "Seuil: $SeuilJours jours (crees avant le $($dateSeuil.ToString('dd/MM/yyyy')))"
Log "Total comptes lus: $cTotalLus"
Log "  Domaines autorises: $cDomainOk"
Log "    Actifs avec acces: $cActifAvecAcces"
Log "    Non actifs: $cNonActif"
Log "    No access (candidats): $cNoAccess"
Log "      Eligibles (>=$SeuilJours jours): $cEligible"
Log "      Non eligibles (<$SeuilJours jours): $cNonEligible"
Log "      Date inconnue: $cDateInconnue"
if ($mode -eq "EXECUTE") {
  Log "      Desactives: $cDisabled"
  Log "      Erreurs: $cError"
}
Log "  CMDB Assets:"
Log "    Trouves: $cAssetTrouve"
Log "    Non trouves: $cAssetNonTrouve"
Log "============================================"
$totalMin = [Math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
Log "Duree totale: $totalMin minutes"
Log "Termine."

# -------------------- Popup resume --------------------
$summaryMsg =
  "Mode: $mode`n" +
  "Seuil: $SeuilJours jours (crees avant le $($dateSeuil.ToString('dd/MM/yyyy')))`n" +
  "Domaines: $($domainesAutorises -join ', ')`n`n" +
  "Total comptes lus: $cTotalLus`n" +
  "Domaines OK: $cDomainOk`n" +
  "  Actifs avec acces: $cActifAvecAcces`n" +
  "  Non actifs: $cNonActif`n" +
  "  No access (candidats): $cNoAccess`n`n" +
  "Eligibles: $cEligible`n" +
  "Non eligibles: $cNonEligible`n" +
  "Date inconnue: $cDateInconnue`n"

if ($mode -eq "EXECUTE") {
  $summaryMsg += "`nDesactives: $cDisabled`nErreurs: $cError`n"
}

$summaryMsg +=
  "`nCMDB Assets: $cAssetTrouve trouves, $cAssetNonTrouve non trouves`n" +
  "`nDuree: $totalMin min`nExport: $outCsv"

[System.Windows.Forms.MessageBox]::Show(
  $summaryMsg,
  "Resultat - Desactivation No Access",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
