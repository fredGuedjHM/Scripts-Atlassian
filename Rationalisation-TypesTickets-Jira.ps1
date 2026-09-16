<#
.SYNOPSIS
  Nettoyage-TypesTickets-Jira.ps1
  Supprime les types inutilises, fusionne les homonymes,
  nettoie les types projet-scoped vides, et identifie
  les projets Team-managed a convertir.

.DESCRIPTION
  Ce script :
  1. Liste tous les types de tickets et compte leur utilisation
  2. Phase A : Supprime les types globaux avec 0 issue
  3. Phase B : Fusionne les homonymes globaux
  4. Phase C : Supprime les types projet-scoped avec 0 issue
  5. Phase D : Rapport des projets Team-managed a convertir

  MODES :
    (defaut)     : Dry-run, affiche le plan sans rien modifier
    -Supprimer   : Phase A uniquement
    -Fusionner   : Phase B uniquement
    -Nettoyer    : Phase C uniquement
    -Rapport     : Phase D uniquement (toujours dry-run)
    -Execute     : Phases A + B + C + D

  SECURITE :
    - Confirmation interactive separee pour chaque phase
    - Log complet de toutes les actions
    - CSV mis a jour avec les statuts

  API :
    GET    /rest/api/3/issuetype
    POST   /rest/api/3/search/jql (nextPageToken dans le body)
    GET    /rest/api/3/issuetype/{id}/alternatives
    DELETE /rest/api/3/issuetype/{id}
    PUT    /rest/api/3/issue/{key}
    GET    /rest/api/3/project/{id}

  CREDENTIALS :
    secrets\site-admin.xml => Jiradot

  FICHIERS GENERES (dans exports\) :
    PlanAction_Types_{ts}.csv  => Plan d action detaille
    PlanAction_Types_{ts}.log  => Journal d execution

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Nettoyage-TypesTickets-Jira.ps1              # Dry-run
  .\Nettoyage-TypesTickets-Jira.ps1 -Supprimer   # Phase A
  .\Nettoyage-TypesTickets-Jira.ps1 -Fusionner   # Phase B
  .\Nettoyage-TypesTickets-Jira.ps1 -Nettoyer    # Phase C
  .\Nettoyage-TypesTickets-Jira.ps1 -Rapport     # Phase D
  .\Nettoyage-TypesTickets-Jira.ps1 -Execute     # Tout
#>

[CmdletBinding()]
param(
  [switch] $Supprimer,
  [switch] $Fusionner,
  [switch] $Nettoyer,
  [switch] $Rapport,
  [switch] $Execute,
  [int]    $MaxRetries = 5,
  [int]    $MaxCount   = 10000
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Determiner le mode
$doSupprimer = $Supprimer -or $Execute
$doFusionner = $Fusionner -or $Execute
$doNettoyer  = $Nettoyer  -or $Execute
$doRapport   = $Rapport   -or $Execute
$isDryRun    = -not $doSupprimer -and -not $doFusionner -and -not $doNettoyer -and -not $doRapport

if ($isDryRun) {
  $modeLabel = "DRY-RUN"
} elseif ($Execute) {
  $modeLabel = "EXECUTION COMPLETE"
} else {
  $parts = @()
  if ($doSupprimer) { $parts += "SUPPRESSION" }
  if ($doFusionner) { $parts += "FUSION" }
  if ($doNettoyer)  { $parts += "NETTOYAGE PROJET" }
  if ($doRapport)   { $parts += "RAPPORT" }
  $modeLabel = $parts -join " + "
}

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("PlanAction_Types_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("PlanAction_Types_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  if ($script:logFile) { Add-Content -Path $script:logFile -Value $line -Encoding UTF8 }
}

function Write-CsvToFile([string]$filePath, [string]$line, [switch]$Header) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  if ($Header) {
    [System.IO.File]::WriteAllText($filePath, "$line`r`n", $utf8Bom)
  } else {
    [System.IO.File]::AppendAllText($filePath, "$line`r`n", $utf8Bom)
  }
}

function CsvEscape([string]$val) {
  return '"{0}"' -f ([string]$val -replace '"','""')
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
      if ($Body) {
        $params["ContentType"] = "application/json; charset=utf-8"
        $params["Body"] = [System.Text.Encoding]::UTF8.GetBytes($Body)
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
      "Credentials $SiteName",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl = Read-Host "URL du site (ex: $SiteName.atlassian.net)"
    $inputUrl = $inputUrl -replace "^https?://", "" -replace "/.*$", "" -replace "/$", ""
    $adminEmail = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $CredFile
    Log "Credentials sauvegardes dans $CredFile"
  }
  $data  = Import-Clixml -Path $CredFile
  $url   = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
    Name    = $SiteName
  }
}

# ============================================================
# NORMALISATION
# ============================================================

function Remove-Diacritics([string]$text) {
  $normalized = $text.Normalize([System.Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($c in $normalized.ToCharArray()) {
    $uc = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
    if ($uc -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($c)
    }
  }
  return $sb.ToString()
}

function Normalize-TypeName([string]$name) {
  $n = Remove-Diacritics $name
  $n = $n.ToLower().Trim()
  $n = $n -replace '[_\-\s\.]+', ' '
  $n = $n -replace '\s+', ' '
  if ($n.EndsWith("s") -and $n.Length -gt 3) { $n = $n.Substring(0, $n.Length - 1) }
  return $n
}

# ============================================================
# COMPTAGE PAR TYPE
# ============================================================

function Count-IssuesByTypeId {
  param([string]$BaseUrl, [hashtable]$Headers, [string]$TypeId, [int]$Cap = 10000)

  $searchUrl = "{0}/rest/api/3/search/jql" -f $BaseUrl
  $nextPageToken = $null
  $count = 0
  $pageSize = 100

  while ($true) {
    $bodyObj = @{
      jql        = "issuetype = $TypeId"
      maxResults = $pageSize
      fields     = @("key")
    }
    if ($nextPageToken) { $bodyObj["nextPageToken"] = $nextPageToken }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress

    $resp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $Headers -Body $bodyJson
    if (-not $resp.ok) { return $count }

    $json = $resp.content | ConvertFrom-Json
    $ic = ($json.issues | Measure-Object).Count
    if ($ic -eq 0) { break }

    $count += $ic
    if ($count -ge $Cap) { return $Cap }
    if ($json.isLast -eq $true) { break }
    if ($json.nextPageToken) { $nextPageToken = [string]$json.nextPageToken } else { break }

    Start-Sleep -Milliseconds 100
  }
  return $count
}

# ============================================================
# MIGRATION INDIVIDUELLE (fallback)
# ============================================================

function Migrate-Issues {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$SourceTypeId,
    [string]$TargetTypeId,
    [string]$SourceTypeName,
    [string]$TargetTypeName
  )

  $searchUrl = "{0}/rest/api/3/search/jql" -f $BaseUrl
  $migrated = 0; $errors = 0

  while ($true) {
    $bodyObj = @{
      jql        = "issuetype = $SourceTypeId ORDER BY key ASC"
      maxResults = 50
      fields     = @("key")
    }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress

    $resp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $Headers -Body $bodyJson
    if (-not $resp.ok) {
      Log ("    Erreur recherche issues type {0} : status={1}" -f $SourceTypeId, $resp.status) "ERROR"
      break
    }

    $json = $resp.content | ConvertFrom-Json
    $issues = $json.issues
    $ic = ($issues | Measure-Object).Count
    if ($ic -eq 0) { break }

    foreach ($issue in $issues) {
      $key = [string]$issue.key
      $updateBody = @{
        fields = @{
          issuetype = @{ id = $TargetTypeId }
        }
      } | ConvertTo-Json -Depth 5 -Compress

      $updateUrl = "{0}/rest/api/3/issue/{1}" -f $BaseUrl, $key
      $updateResp = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $Headers -Body $updateBody

      if ($updateResp.ok) {
        $migrated++
      } else {
        $errors++
        Log ("    ERREUR migration {0} : status={1}" -f $key, $updateResp.status) "ERROR"
      }

      if (($migrated + $errors) % 50 -eq 0) {
        Write-Progress -Activity "Migration $SourceTypeName -> $TargetTypeName" `
          -Status ("{0} migres, {1} erreurs" -f $migrated, $errors) -PercentComplete (-1)
      }
    }

    Start-Sleep -Milliseconds 200
  }

  Write-Progress -Activity "Migration" -Completed
  return @{ Migrated=$migrated; Errors=$errors }
}

# ============================================================
# DEBUT
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  NETTOYAGE DES TYPES DE TICKETS" -ForegroundColor Cyan
Write-Host ("  Mode : {0}" -f $modeLabel) -ForegroundColor $(if ($isDryRun) { "Green" } else { "Red" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($isDryRun) {
  Write-Host "  Mode DRY-RUN : aucune modification ne sera effectuee." -ForegroundColor Green
  Write-Host "  Options disponibles :" -ForegroundColor DarkGray
  Write-Host "    -Supprimer  : Phase A - supprime les types globaux avec 0 issue" -ForegroundColor DarkGray
  Write-Host "    -Fusionner  : Phase B - fusionne les homonymes globaux" -ForegroundColor DarkGray
  Write-Host "    -Nettoyer   : Phase C - supprime les types projet-scoped avec 0 issue" -ForegroundColor DarkGray
  Write-Host "    -Rapport    : Phase D - rapport des projets Team-managed a convertir" -ForegroundColor DarkGray
  Write-Host "    -Execute    : toutes les phases" -ForegroundColor DarkGray
  Write-Host ""
}

Log "================================================================"
Log ("  NETTOYAGE DES TYPES DE TICKETS - {0}" -f $modeLabel)
Log "================================================================"

$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
Log ("  Site : {0}" -f $site.BaseUrl)

$startTime = Get-Date

# ============================================================
# ETAPE 1 : LISTING DES TYPES
# ============================================================

Log "=== ETAPE 1 : Listing des types ==="

$url = "{0}/rest/api/3/issuetype" -f $site.BaseUrl
$resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

if (-not $resp.ok) {
  Log ("  Erreur listing : status={0}" -f $resp.status) "ERROR"
  return
}

$issueTypes = $resp.content | ConvertFrom-Json
$cTypes = ($issueTypes | Measure-Object).Count
Log ("  {0} types trouves" -f $cTypes)

$typeData = New-Object System.Collections.Generic.List[object]
foreach ($it in $issueTypes) {
  $scope = ""
  $scopeProjectId = ""
  if ($it.scope) {
    if ($it.scope.type -eq "PROJECT") {
      $scope = "PROJECT"
      $scopeProjectId = [string]$it.scope.project.id
    } else {
      $scope = [string]$it.scope.type
    }
  }
  $typeData.Add(@{
    Id             = [string]$it.id
    Name           = [string]$it.name
    IsSubtask      = [bool]$it.subtask
    Scope          = $scope
    ScopeProjectId = $scopeProjectId
    Normalized     = Normalize-TypeName ([string]$it.name)
    IssueCount     = -1
  }) | Out-Null
}

Write-Host ("  {0} types trouves" -f $cTypes) -ForegroundColor Green

# ============================================================
# ETAPE 2 : COMPTAGE PAR ID
# ============================================================

Log "=== ETAPE 2 : Comptage par type ==="

$cDone = 0
foreach ($t in $typeData) {
  $cDone++
  Write-Progress -Activity "Comptage" `
    -Status ("{0}/{1} : {2} (ID:{3})" -f $cDone, $cTypes, $t.Name, $t.Id) `
    -PercentComplete ([int](100 * $cDone / $cTypes))

  $count = Count-IssuesByTypeId -BaseUrl $site.BaseUrl -Headers $site.Headers -TypeId $t.Id -Cap $MaxCount
  $t.IssueCount = $count

  $scopeLabel = if ($t.Scope -eq "PROJECT") { " [Projet:{0}]" -f $t.ScopeProjectId } else { " [Global]" }
  $countLabel = if ($count -ge $MaxCount) { "{0}+" -f $count } else { "$count" }
  Log ("  {0,-35} (ID:{1}){2} : {3}" -f $t.Name, $t.Id, $scopeLabel, $countLabel)
}
Write-Progress -Activity "Comptage" -Completed

# ============================================================
# ETAPE 3 : GROUPES D'HOMONYMES
# ============================================================

Log "=== ETAPE 3 : Groupes d'homonymes ==="

$globalTypes  = $typeData | Where-Object { $_.Scope -ne "PROJECT" }
$projectTypes = $typeData | Where-Object { $_.Scope -eq "PROJECT" }

$groups = @{}
foreach ($t in $globalTypes) {
  $norm = $t.Normalized
  if (-not $groups.ContainsKey($norm)) { $groups[$norm] = @() }
  $groups[$norm] += $t
}

$fusionGroups = @{}
foreach ($k in $groups.Keys) {
  if ($groups[$k].Count -gt 1) {
    $fusionGroups[$k] = $groups[$k]
  }
}

$cFusionGroups = $fusionGroups.Count
Log ("  {0} groupes d'homonymes globaux" -f $cFusionGroups)

# ============================================================
# ETAPE 3b : ANALYSE DES PROJETS TEAM-MANAGED
# ============================================================

Log "=== ETAPE 3b : Analyse des projets Team-managed ==="

# Regrouper les types projet-scoped par projet
$projectTypeGroups = @{}
foreach ($t in $projectTypes) {
  $pid = $t.ScopeProjectId
  if (-not $projectTypeGroups.ContainsKey($pid)) {
    $projectTypeGroups[$pid] = @()
  }
  $projectTypeGroups[$pid] += $t
}

# Recuperer les infos de chaque projet concerne
$projectInfos = @{}
foreach ($pid in $projectTypeGroups.Keys) {
  $projUrl = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $pid
  $projResp = Invoke-ApiCall -Method "GET" -Url $projUrl -Headers $site.Headers
  if ($projResp.ok) {
    $projJson = $projResp.content | ConvertFrom-Json
    $projectInfos[$pid] = @{
      Key   = [string]$projJson.key
      Name  = [string]$projJson.name
      Style = if ($projJson.style) { [string]$projJson.style } else { "classic" }
      Lead  = if ($projJson.lead -and $projJson.lead.displayName) { [string]$projJson.lead.displayName } else { "" }
    }
    Log ("  Projet {0} ({1}) : style={2}, lead={3}" -f $projJson.key, $projJson.name, $projectInfos[$pid].Style, $projectInfos[$pid].Lead)
  } else {
    $projectInfos[$pid] = @{ Key="?"; Name="Projet ID $pid"; Style="inconnu"; Lead="" }
    Log ("  Projet ID {0} : inaccessible (status={1})" -f $pid, $projResp.status) "WARN"
  }
  Start-Sleep -Milliseconds 100
}

$cTeamManaged = ($projectTypeGroups.Keys | Measure-Object).Count
Log ("  {0} projets avec types projet-scoped" -f $cTeamManaged)

# ============================================================
# ETAPE 4 : PLAN D'ACTION
# ============================================================

Log "=== ETAPE 4 : Plan d'action ==="

$plan = New-Object System.Collections.Generic.List[object]

# IDs des types dans les groupes de fusion
$typesInFusionGroups = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $fusionGroups.Keys) {
  foreach ($t in $fusionGroups[$k]) {
    [void]$typesInFusionGroups.Add($t.Id)
  }
}

# --- Phase A : Suppressions globales (0 issues, hors groupes de fusion) ---
foreach ($t in $globalTypes) {
  if ($t.IssueCount -eq 0 -and -not $typesInFusionGroups.Contains($t.Id)) {
    $plan.Add(@{
      Phase      = "A-SUPPRIMER"
      Action     = "SUPPRIMER"
      SourceName = $t.Name
      SourceId   = $t.Id
      SourceScope= "Global"
      SourceCount= 0
      TargetName = ""
      TargetId   = ""
      ProjectKey = ""
      ProjectName= ""
      Reason     = "0 issue, type global inutilise"
      Status     = "PLANIFIE"
    }) | Out-Null
  }
}

# --- Phase B : Fusions (homonymes globaux) ---
foreach ($k in ($fusionGroups.Keys | Sort-Object)) {
  $members = $fusionGroups[$k] | Sort-Object { $_.IssueCount } -Descending
  $target  = $members[0]
  $sources = $members | Select-Object -Skip 1

  foreach ($src in $sources) {
    if ($src.IssueCount -eq 0) {
      $reasonText = "Homonyme de '{0}' (ID:{1}), 0 issue" -f $target.Name, $target.Id
      $plan.Add(@{
        Phase      = "B-FUSIONNER"
        Action     = "SUPPRIMER"
        SourceName = $src.Name
        SourceId   = $src.Id
        SourceScope= "Global"
        SourceCount= 0
        TargetName = $target.Name
        TargetId   = $target.Id
        ProjectKey = ""
        ProjectName= ""
        Reason     = $reasonText
        Status     = "PLANIFIE"
      }) | Out-Null
    } else {
      $reasonText = "Homonyme de '{0}' (ID:{1}, {2} issues), migrer {3} issues" -f $target.Name, $target.Id, $target.IssueCount, $src.IssueCount
      $plan.Add(@{
        Phase      = "B-FUSIONNER"
        Action     = "MIGRER+SUPPRIMER"
        SourceName = $src.Name
        SourceId   = $src.Id
        SourceScope= "Global"
        SourceCount= $src.IssueCount
        TargetName = $target.Name
        TargetId   = $target.Id
        ProjectKey = ""
        ProjectName= ""
        Reason     = $reasonText
        Status     = "PLANIFIE"
      }) | Out-Null
    }
  }
}

# --- Phase C : Types projet-scoped avec 0 issue ---
foreach ($t in $projectTypes) {
  if ($t.IssueCount -eq 0) {
    $pInfo = $projectInfos[$t.ScopeProjectId]
    $projKey  = if ($pInfo) { $pInfo.Key } else { "?" }
    $projName = if ($pInfo) { $pInfo.Name } else { "?" }
    $scopeText = "project:{0}" -f $t.ScopeProjectId
    $reasonText = "Type projet-scoped ({0}), 0 issue" -f $projKey
    $plan.Add(@{
      Phase      = "C-NETTOYER"
      Action     = "SUPPRIMER"
      SourceName = $t.Name
      SourceId   = $t.Id
      SourceScope= $scopeText
      SourceCount= 0
      TargetName = ""
      TargetId   = ""
      ProjectKey = $projKey
      ProjectName= $projName
      Reason     = $reasonText
      Status     = "PLANIFIE"
    }) | Out-Null
  }
}

# --- Phase D : Rapport projets Team-managed ---
foreach ($pid in ($projectTypeGroups.Keys | Sort-Object)) {
  $pTypes = $projectTypeGroups[$pid]
  $pInfo  = $projectInfos[$pid]
  $projKey  = if ($pInfo) { $pInfo.Key } else { "?" }
  $projName = if ($pInfo) { $pInfo.Name } else { "?" }
  $projStyle = if ($pInfo) { $pInfo.Style } else { "?" }
  $projLead  = if ($pInfo) { $pInfo.Lead } else { "?" }

  $cTypesProj  = ($pTypes | Measure-Object).Count
  $cTypesVides = ($pTypes | Where-Object { $_.IssueCount -eq 0 } | Measure-Object).Count
  $cTypesActifs= $cTypesProj - $cTypesVides
  $typeNames   = ($pTypes | ForEach-Object { $_.Name } | Sort-Object -Unique) -join ", "
  $totalIssues = ($pTypes | Measure-Object -Property IssueCount -Sum).Sum

  $reasonText = "Projet {0} ({1}), style={2}, lead={3}, {4} types ({5} vides, {6} actifs), {7} issues, types: {8}" -f `
    $projKey, $projName, $projStyle, $projLead, $cTypesProj, $cTypesVides, $cTypesActifs, $totalIssues, $typeNames

  $actionText = if ($cTypesActifs -eq 0) { "CONVERTIR (aucune issue)" } else { "CONVERTIR (migration requise)" }

  $plan.Add(@{
    Phase      = "D-RAPPORT"
    Action     = $actionText
    SourceName = $projKey
    SourceId   = $pid
    SourceScope= "Team-managed"
    SourceCount= $totalIssues
    TargetName = ""
    TargetId   = ""
    ProjectKey = $projKey
    ProjectName= $projName
    Reason     = $reasonText
    Status     = "A EVALUER"
  }) | Out-Null
}

$cPhaseA    = ($plan | Where-Object { $_.Phase -eq "A-SUPPRIMER" } | Measure-Object).Count
$cPhaseBSup = ($plan | Where-Object { $_.Phase -eq "B-FUSIONNER" -and $_.Action -eq "SUPPRIMER" } | Measure-Object).Count
$cPhaseBMig = ($plan | Where-Object { $_.Phase -eq "B-FUSIONNER" -and $_.Action -eq "MIGRER+SUPPRIMER" } | Measure-Object).Count
$cPhaseC    = ($plan | Where-Object { $_.Phase -eq "C-NETTOYER" } | Measure-Object).Count
$cPhaseD    = ($plan | Where-Object { $_.Phase -eq "D-RAPPORT" } | Measure-Object).Count

Log ("  Phase A : {0} suppressions globales" -f $cPhaseA)
Log ("  Phase B : {0} suppressions homonymes + {1} fusions" -f $cPhaseBSup, $cPhaseBMig)
Log ("  Phase C : {0} suppressions projet-scoped" -f $cPhaseC)
Log ("  Phase D : {0} projets Team-managed a evaluer" -f $cPhaseD)

# ============================================================
# ETAPE 5 : EXPORT CSV DU PLAN
# ============================================================

Log "=== ETAPE 5 : Export du plan ==="

$csvCols = @("Phase","Action","SourceName","SourceId","SourceScope","SourceCount","TargetName","TargetId","ProjectKey","ProjectName","Reason","Status")
$csvHeaderLine = ($csvCols | ForEach-Object { CsvEscape $_ }) -join ";"
Write-CsvToFile $csvFile $csvHeaderLine -Header

foreach ($p in $plan) {
  $vals = @(
    $p.Phase, $p.Action, $p.SourceName, $p.SourceId, $p.SourceScope,
    [string]$p.SourceCount, $p.TargetName, $p.TargetId, $p.ProjectKey, $p.ProjectName, $p.Reason, $p.Status
  )
  $line = ($vals | ForEach-Object { CsvEscape $_ }) -join ";"
  Write-CsvToFile $csvFile $line
}

Log ("  Plan CSV -> {0}" -f $csvFile)

# ============================================================
# ETAPE 6 : AFFICHAGE DU PLAN
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  PLAN D'ACTION ({0})" -f $modeLabel) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Phase A ---
$phaseAItems = $plan | Where-Object { $_.Phase -eq "A-SUPPRIMER" }
Write-Host ""
Write-Host ("  PHASE A - SUPPRESSIONS GLOBALES ({0} types avec 0 issue) :" -f $cPhaseA) -ForegroundColor Red
if ($cPhaseA -eq 0) {
  Write-Host "    Aucun type global a supprimer." -ForegroundColor DarkGray
} else {
  foreach ($p in ($phaseAItems | Sort-Object { $_.SourceName })) {
    Write-Host ("    SUPPRIMER  {0,-35} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor Red
  }
}

# --- Phase B ---
$phaseBItems = $plan | Where-Object { $_.Phase -eq "B-FUSIONNER" }
$cPhaseB = ($phaseBItems | Measure-Object).Count
Write-Host ""
Write-Host ("  PHASE B - FUSIONS HOMONYMES ({0} actions) :" -f $cPhaseB) -ForegroundColor DarkYellow
if ($cPhaseB -eq 0) {
  Write-Host "    Aucune fusion necessaire." -ForegroundColor DarkGray
} else {
  foreach ($p in ($phaseBItems | Sort-Object { $_.SourceName })) {
    if ($p.Action -eq "MIGRER+SUPPRIMER") {
      Write-Host ("    MIGRER     {0,-30} (ID:{1}, {2} issues) -> {3} (ID:{4})" -f $p.SourceName, $p.SourceId, $p.SourceCount, $p.TargetName, $p.TargetId) -ForegroundColor DarkYellow
    } else {
      Write-Host ("    SUPPRIMER  {0,-30} (ID:{1}, 0 issue) [homonyme de {2}]" -f $p.SourceName, $p.SourceId, $p.TargetName) -ForegroundColor Red
    }
  }
}

# --- Phase C ---
$phaseCItems = $plan | Where-Object { $_.Phase -eq "C-NETTOYER" }
Write-Host ""
Write-Host ("  PHASE C - TYPES PROJET-SCOPED VIDES ({0} types) :" -f $cPhaseC) -ForegroundColor Magenta
if ($cPhaseC -eq 0) {
  Write-Host "    Aucun type projet-scoped a supprimer." -ForegroundColor DarkGray
} else {
  # Grouper par projet pour la lisibilite
  $phaseCByProject = $phaseCItems | Group-Object { $_.ProjectKey } | Sort-Object Name
  foreach ($grp in $phaseCByProject) {
    $projKey = $grp.Name
    $projName = $grp.Group[0].ProjectName
    $grpCount = ($grp.Group | Measure-Object).Count
    Write-Host ("    Projet {0} ({1}) - {2} types :" -f $projKey, $projName, $grpCount) -ForegroundColor Magenta
    foreach ($p in ($grp.Group | Sort-Object { $_.SourceName })) {
      Write-Host ("      SUPPRIMER  {0,-30} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor DarkMagenta
    }
  }
}

# --- Phase D ---
$phaseDItems = $plan | Where-Object { $_.Phase -eq "D-RAPPORT" }
Write-Host ""
Write-Host ("  PHASE D - PROJETS TEAM-MANAGED A CONVERTIR ({0} projets) :" -f $cPhaseD) -ForegroundColor Blue
if ($cPhaseD -eq 0) {
  Write-Host "    Aucun projet Team-managed detecte." -ForegroundColor DarkGray
} else {
  foreach ($p in ($phaseDItems | Sort-Object { $_.SourceName })) {
    $pTypes = $projectTypeGroups[$p.SourceId]
    $cTypesProj  = ($pTypes | Measure-Object).Count
    $cTypesVides = ($pTypes | Where-Object { $_.IssueCount -eq 0 } | Measure-Object).Count
    $color = if ($p.SourceCount -eq 0) { "Green" } else { "Yellow" }
    Write-Host ("    {0,-10} {1,-30} {2,5} issues, {3} types ({4} vides)  [{5}]" -f `
      $p.SourceName, $p.ProjectName, $p.SourceCount, $cTypesProj, $cTypesVides, $p.Action) -ForegroundColor $color
  }
  Write-Host ""
  Write-Host "    Pour convertir : Parametres du projet > Fonctionnalites > Convertir en Company-managed" -ForegroundColor DarkGray
}

Write-Host ""

# ============================================================
# ETAPE 7 : EXECUTION
# ============================================================

if ($isDryRun) {
  Write-Host "========================================" -ForegroundColor Green
  Write-Host "  DRY-RUN TERMINE" -ForegroundColor Green
  Write-Host "  Aucune modification effectuee." -ForegroundColor Green
  Write-Host "" -ForegroundColor Green
  Write-Host "  Pour executer :" -ForegroundColor DarkGray
  Write-Host "    -Supprimer  : Phase A" -ForegroundColor DarkGray
  Write-Host "    -Fusionner  : Phase B" -ForegroundColor DarkGray
  Write-Host "    -Nettoyer   : Phase C" -ForegroundColor DarkGray
  Write-Host "    -Rapport    : Phase D (info seulement)" -ForegroundColor DarkGray
  Write-Host "    -Execute    : tout" -ForegroundColor DarkGray
  Write-Host ""
  Write-Host ("  Plan CSV : {0}" -f $csvFile)
  Write-Host ("  LOG      : {0}" -f $logFile)
  Write-Host "========================================" -ForegroundColor Green
  Log "DRY-RUN termine."
  return
}

$cOK = 0; $cKO = 0; $cSkip = 0

# ============================================================
# PHASE A : SUPPRESSIONS GLOBALES
# ============================================================

if ($doSupprimer -and $cPhaseA -gt 0) {
  Write-Host "========================================" -ForegroundColor Red
  Write-Host "  PHASE A : SUPPRESSION DES TYPES GLOBAUX VIDES" -ForegroundColor Red
  Write-Host "========================================" -ForegroundColor Red
  Write-Host ""
  Write-Host ("  {0} types avec 0 issue vont etre supprimes." -f $cPhaseA) -ForegroundColor Red
  Write-Host "  Cette operation est IRREVERSIBLE." -ForegroundColor Red
  Write-Host ""
  $confirmA = Read-Host "  Tapez SUPPRIMER pour confirmer (ou autre chose pour passer)"

  if ($confirmA -eq "SUPPRIMER") {
    Log "=== PHASE A : Execution ==="

    foreach ($p in $phaseAItems) {
      Log ("  Suppression '{0}' (ID:{1})..." -f $p.SourceName, $p.SourceId)

      $delUrl = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $p.SourceId
      $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers

      if ($delResp.ok -or $delResp.status -eq 204) {
        $p.Status = "OK"
        $cOK++
        Log ("    OK : '{0}' supprime" -f $p.SourceName)
        Write-Host ("    OK  {0,-35} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor Green
      } else {
        $errStatus = "ERREUR:{0}" -f $delResp.status
        $p.Status = $errStatus
        $cKO++
        Log ("    ERREUR : status={0}" -f $delResp.status) "ERROR"
        Write-Host ("    ERR {0,-35} (ID:{1}) status={2}" -f $p.SourceName, $p.SourceId, $delResp.status) -ForegroundColor Red

        if ($delResp.body) {
          try {
            $errJson = $delResp.body | ConvertFrom-Json
            if ($errJson.errorMessages) {
              foreach ($em in $errJson.errorMessages) {
                Log ("      Detail : {0}" -f $em) "ERROR"
                Write-Host ("      -> {0}" -f $em) -ForegroundColor DarkGray
              }
            }
          } catch {}
        }
      }

      Start-Sleep -Milliseconds 300
    }
    Write-Host ""
  } else {
    Log "  Phase A ignoree par l'utilisateur."
    Write-Host "  Phase A ignoree." -ForegroundColor Yellow
    Write-Host ""
    foreach ($p in $phaseAItems) { $p.Status = "IGNORE"; $cSkip++ }
  }
} elseif ($doSupprimer) {
  Write-Host "  Phase A : aucun type global a supprimer." -ForegroundColor DarkGray
  Write-Host ""
}

# ============================================================
# PHASE B : FUSIONS
# ============================================================

if ($doFusionner -and $cPhaseB -gt 0) {
  Write-Host "========================================" -ForegroundColor DarkYellow
  Write-Host "  PHASE B : FUSION DES HOMONYMES" -ForegroundColor DarkYellow
  Write-Host "========================================" -ForegroundColor DarkYellow
  Write-Host ""
  Write-Host ("  {0} suppressions d'homonymes vides + {1} fusions avec migration." -f $cPhaseBSup, $cPhaseBMig) -ForegroundColor DarkYellow
  if ($cPhaseBMig -gt 0) {
    Write-Host "  Les issues seront migrees vers le type cible avant suppression." -ForegroundColor DarkYellow
  }
  Write-Host "  Cette operation est IRREVERSIBLE." -ForegroundColor Red
  Write-Host ""
  $confirmB = Read-Host "  Tapez FUSIONNER pour confirmer (ou autre chose pour passer)"

  if ($confirmB -eq "FUSIONNER") {
    Log "=== PHASE B : Execution ==="

    foreach ($p in $phaseBItems) {

      if ($p.Action -eq "SUPPRIMER") {
        Log ("  Suppression homonyme '{0}' (ID:{1})..." -f $p.SourceName, $p.SourceId)

        $delUrl = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $p.SourceId
        $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers

        if ($delResp.ok -or $delResp.status -eq 204) {
          $p.Status = "OK"
          $cOK++
          Log ("    OK : homonyme '{0}' supprime" -f $p.SourceName)
          Write-Host ("    OK  SUPPR  {0,-30} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor Green
        } else {
          $errStatus = "ERREUR:{0}" -f $delResp.status
          $p.Status = $errStatus
          $cKO++
          Log ("    ERREUR : status={0}" -f $delResp.status) "ERROR"
          Write-Host ("    ERR SUPPR  {0,-30} (ID:{1}) status={2}" -f $p.SourceName, $p.SourceId, $delResp.status) -ForegroundColor Red
        }

      } elseif ($p.Action -eq "MIGRER+SUPPRIMER") {
        $logMsg = "  Fusion '{0}' (ID:{1}, {2} issues) -> '{3}' (ID:{4})..." -f $p.SourceName, $p.SourceId, $p.SourceCount, $p.TargetName, $p.TargetId
        Log $logMsg
        Write-Host ("    FUSION {0,-25} (ID:{1}, {2}) -> {3} (ID:{4})" -f $p.SourceName, $p.SourceId, $p.SourceCount, $p.TargetName, $p.TargetId) -ForegroundColor DarkYellow

        # Verifier les alternatives
        $altUrl = "{0}/rest/api/3/issuetype/{1}/alternatives" -f $site.BaseUrl, $p.SourceId
        $altResp = Invoke-ApiCall -Method "GET" -Url $altUrl -Headers $site.Headers
        $altOk = $false

        if ($altResp.ok) {
          $alternatives = $altResp.content | ConvertFrom-Json
          foreach ($alt in $alternatives) {
            if ([string]$alt.id -eq $p.TargetId) { $altOk = $true; break }
          }
        }

        if ($altOk) {
          $delUrl = "{0}/rest/api/3/issuetype/{1}?alternativeIssueTypeId={2}" -f $site.BaseUrl, $p.SourceId, $p.TargetId
          $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers

          if ($delResp.ok -or $delResp.status -eq 204) {
            $p.Status = "OK"
            $cOK++
            Log ("    OK : {0} issues migrees, type supprime" -f $p.SourceCount)
            Write-Host ("    OK  {0} issues migrees, type supprime" -f $p.SourceCount) -ForegroundColor Green
          } else {
            Log ("    DELETE echoue (status={0}), fallback..." -f $delResp.status) "WARN"
            $migResult = Migrate-Issues -BaseUrl $site.BaseUrl -Headers $site.Headers `
              -SourceTypeId $p.SourceId -TargetTypeId $p.TargetId `
              -SourceTypeName $p.SourceName -TargetTypeName $p.TargetName

            if ($migResult.Errors -eq 0 -and $migResult.Migrated -gt 0) {
              $delUrl2 = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $p.SourceId
              $delResp2 = Invoke-ApiCall -Method "DELETE" -Url $delUrl2 -Headers $site.Headers
              if ($delResp2.ok -or $delResp2.status -eq 204) {
                $p.Status = "OK (fallback)"
                $cOK++
                Log ("    OK fallback : {0} migrees, supprime" -f $migResult.Migrated)
                Write-Host ("    OK  {0} migrees (fallback), supprime" -f $migResult.Migrated) -ForegroundColor Green
              } else {
                $p.Status = "PARTIEL"
                $cKO++
                Log "    PARTIEL : migre mais suppression echouee" "WARN"
                Write-Host "    WARN  Migre mais suppression echouee" -ForegroundColor Yellow
              }
            } else {
              $errMsg = "ERREUR ({0} OK, {1} err)" -f $migResult.Migrated, $migResult.Errors
              $p.Status = $errMsg
              $cKO++
              Log "    ERREUR : migration incomplete" "ERROR"
              Write-Host ("    ERR  Migration incomplete ({0} erreurs)" -f $migResult.Errors) -ForegroundColor Red
            }
          }
        } else {
          Log "    Cible non dans alternatives, migration individuelle..." "WARN"
          $migResult = Migrate-Issues -BaseUrl $site.BaseUrl -Headers $site.Headers `
            -SourceTypeId $p.SourceId -TargetTypeId $p.TargetId `
            -SourceTypeName $p.SourceName -TargetTypeName $p.TargetName

          if ($migResult.Errors -eq 0 -and $migResult.Migrated -gt 0) {
            $delUrl3 = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $p.SourceId
            $delResp3 = Invoke-ApiCall -Method "DELETE" -Url $delUrl3 -Headers $site.Headers
            if ($delResp3.ok -or $delResp3.status -eq 204) {
              $p.Status = "OK (migration individuelle)"
              $cOK++
              Log ("    OK : {0} migrees, supprime" -f $migResult.Migrated)
              Write-Host ("    OK  {0} migrees, supprime" -f $migResult.Migrated) -ForegroundColor Green
            } else {
              $p.Status = "PARTIEL"
              $cKO++
              Write-Host "    WARN  Migre mais suppression echouee" -ForegroundColor Yellow
            }
          } elseif ($migResult.Errors -gt 0) {
            $errMsg = "ERREUR ({0} OK, {1} err)" -f $migResult.Migrated, $migResult.Errors
            $p.Status = $errMsg
            $cKO++
            Write-Host ("    ERR  Migration incomplete ({0} erreurs)" -f $migResult.Errors) -ForegroundColor Red
          } else {
            $p.Status = "ERREUR (0 migre)"
            $cKO++
            Write-Host "    ERR  Aucune issue migree" -ForegroundColor Red
          }
        }
      }

      Start-Sleep -Milliseconds 300
    }
    Write-Host ""
  } else {
    Log "  Phase B ignoree par l'utilisateur."
    Write-Host "  Phase B ignoree." -ForegroundColor Yellow
    Write-Host ""
    foreach ($p in $phaseBItems) { $p.Status = "IGNORE"; $cSkip++ }
  }
} elseif ($doFusionner) {
  Write-Host "  Phase B : aucune fusion necessaire." -ForegroundColor DarkGray
  Write-Host ""
}

# ============================================================
# PHASE C : NETTOYAGE TYPES PROJET-SCOPED
# ============================================================

if ($doNettoyer -and $cPhaseC -gt 0) {
  Write-Host "========================================" -ForegroundColor Magenta
  Write-Host "  PHASE C : NETTOYAGE TYPES PROJET-SCOPED VIDES" -ForegroundColor Magenta
  Write-Host "========================================" -ForegroundColor Magenta
  Write-Host ""
  Write-Host ("  {0} types projet-scoped avec 0 issue vont etre supprimes." -f $cPhaseC) -ForegroundColor Magenta
  Write-Host "  Ces types sont lies a des projets Team-managed." -ForegroundColor DarkGray
  Write-Host "  Si la suppression echoue, le type est conserve (sans impact)." -ForegroundColor DarkGray
  Write-Host ""
  $confirmC = Read-Host "  Tapez NETTOYER pour confirmer (ou autre chose pour passer)"

  if ($confirmC -eq "NETTOYER") {
    Log "=== PHASE C : Execution ==="

    foreach ($p in $phaseCItems) {
      $logMsg = "  Suppression '{0}' (ID:{1}) du projet {2}..." -f $p.SourceName, $p.SourceId, $p.ProjectKey
      Log $logMsg

      $delUrl = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $p.SourceId
      $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers

      if ($delResp.ok -or $delResp.status -eq 204) {
        $p.Status = "OK"
        $cOK++
        Log ("    OK : '{0}' supprime (projet {1})" -f $p.SourceName, $p.ProjectKey)
        Write-Host ("    OK  {0,-25} (ID:{1}) [{2}]" -f $p.SourceName, $p.SourceId, $p.ProjectKey) -ForegroundColor Green
      } else {
        $errStatus = "ERREUR:{0}" -f $delResp.status
        $p.Status = $errStatus
        $cKO++
        Log ("    ERREUR : status={0}" -f $delResp.status) "ERROR"
        Write-Host ("    ERR {0,-25} (ID:{1}) [{2}] status={3}" -f $p.SourceName, $p.SourceId, $p.ProjectKey, $delResp.status) -ForegroundColor Red

        if ($delResp.body) {
          try {
            $errJson = $delResp.body | ConvertFrom-Json
            if ($errJson.errorMessages) {
              foreach ($em in $errJson.errorMessages) {
                Log ("      Detail : {0}" -f $em) "ERROR"
                Write-Host ("      -> {0}" -f $em) -ForegroundColor DarkGray
              }
            }
          } catch {}
        }
      }

      Start-Sleep -Milliseconds 300
    }
    Write-Host ""
  } else {
    Log "  Phase C ignoree par l'utilisateur."
    Write-Host "  Phase C ignoree." -ForegroundColor Yellow
    Write-Host ""
    foreach ($p in $phaseCItems) { $p.Status = "IGNORE"; $cSkip++ }
  }
} elseif ($doNettoyer) {
  Write-Host "  Phase C : aucun type projet-scoped a supprimer." -ForegroundColor DarkGray
  Write-Host ""
}

# ============================================================
# PHASE D : RAPPORT PROJETS TEAM-MANAGED
# ============================================================

if ($doRapport -and $cPhaseD -gt 0) {
  Write-Host "========================================" -ForegroundColor Blue
  Write-Host "  PHASE D : RAPPORT PROJETS TEAM-MANAGED" -ForegroundColor Blue
  Write-Host "========================================" -ForegroundColor Blue
  Write-Host ""
  Write-Host "  Ces projets utilisent des types de tickets isoles." -ForegroundColor DarkGray
  Write-Host "  Les convertir en Company-managed eliminera les doublons." -ForegroundColor DarkGray
  Write-Host ""

  foreach ($p in ($phaseDItems | Sort-Object { $_.SourceName })) {
    $pTypes = $projectTypeGroups[$p.SourceId]
    $cTypesProj  = ($pTypes | Measure-Object).Count
    $cTypesVides = ($pTypes | Where-Object { $_.IssueCount -eq 0 } | Measure-Object).Count
    $cTypesActifs= $cTypesProj - $cTypesVides
    $pInfo = $projectInfos[$p.SourceId]
    $projStyle = if ($pInfo) { $pInfo.Style } else { "?" }
    $projLead  = if ($pInfo) { $pInfo.Lead } else { "?" }

    $color = if ($p.SourceCount -eq 0) { "Green" } else { "Yellow" }
    $riskLabel = if ($p.SourceCount -eq 0) { "SANS RISQUE" } else { "MIGRATION REQUISE" }

    Write-Host ("    {0,-10} {1,-30}" -f $p.SourceName, $p.ProjectName) -ForegroundColor $color -NoNewline
    Write-Host (" style={0}" -f $projStyle) -ForegroundColor DarkGray -NoNewline
    Write-Host (" lead={0}" -f $projLead) -ForegroundColor DarkGray
    Write-Host ("               {0} issues, {1} types ({2} vides)  -> {3}" -f $p.SourceCount, $cTypesProj, $cTypesVides, $riskLabel) -ForegroundColor $color

    # Lister les types du projet
    foreach ($pt in ($pTypes | Sort-Object { $_.Name })) {
      $ptCount = if ($pt.IssueCount -ge $MaxCount) { "{0}+" -f $pt.IssueCount } else { [string]$pt.IssueCount }
      $ptColor = if ($pt.IssueCount -eq 0) { "DarkGray" } else { "White" }
      Write-Host ("                 - {0,-25} (ID:{1}) {2} issues" -f $pt.Name, $pt.Id, $ptCount) -ForegroundColor $ptColor
    }
    Write-Host ""
  }

  Write-Host "  PROCEDURE DE CONVERSION :" -ForegroundColor Blue
  Write-Host "    1. Aller dans Parametres du projet" -ForegroundColor DarkGray
  Write-Host "    2. Section 'Fonctionnalites' ou 'Features'" -ForegroundColor DarkGray
  Write-Host "    3. Cliquer 'Convertir en projet gere par l'entreprise'" -ForegroundColor DarkGray
  Write-Host "    4. Suivre l'assistant de migration" -ForegroundColor DarkGray
  Write-Host "    5. Verifier que les types globaux sont bien utilises" -ForegroundColor DarkGray
  Write-Host ""
  Write-Host "  RECOMMANDATION :" -ForegroundColor Blue
  Write-Host "    - Commencer par les projets SANS RISQUE (0 issue)" -ForegroundColor DarkGray
  Write-Host "    - Pour les projets avec issues, planifier avec le lead du projet" -ForegroundColor DarkGray
  Write-Host ""

  # Marquer comme traite
  foreach ($p in $phaseDItems) { $p.Status = "RAPPORT GENERE" }
} elseif ($doRapport) {
  Write-Host "  Phase D : aucun projet Team-managed detecte." -ForegroundColor DarkGray
  Write-Host ""
}

# ============================================================
# MISE A JOUR DU CSV AVEC LES STATUTS
# ============================================================

Log "  Mise a jour du CSV avec les statuts..."
Write-CsvToFile $csvFile $csvHeaderLine -Header
foreach ($p in $plan) {
  $vals = @(
    $p.Phase, $p.Action, $p.SourceName, $p.SourceId, $p.SourceScope,
    [string]$p.SourceCount, $p.TargetName, $p.TargetId, $p.ProjectKey, $p.ProjectName, $p.Reason, $p.Status
  )
  $line = ($vals | ForEach-Object { CsvEscape $_ }) -join ";"
  Write-CsvToFile $csvFile $line
}

# ============================================================
# RESUME FINAL
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  RESUME FINAL ({0})" -f $modeLabel) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Types analyses        : {0}" -f $cTypes)
Write-Host ("    - Globaux           : {0}" -f ($globalTypes | Measure-Object).Count)
Write-Host ("    - Projet-scoped     : {0}" -f ($projectTypes | Measure-Object).Count)
Write-Host ""
Write-Host ("  Phase A planifiees    : {0} suppressions globales" -f $cPhaseA)
Write-Host ("  Phase B planifiees    : {0} fusions ({1} suppr + {2} migr)" -f ($cPhaseBSup + $cPhaseBMig), $cPhaseBSup, $cPhaseBMig)
Write-Host ("  Phase C planifiees    : {0} suppressions projet-scoped" -f $cPhaseC)
Write-Host ("  Phase D               : {0} projets Team-managed" -f $cPhaseD)
Write-Host ""

if (-not $isDryRun) {
  Write-Host ("  Reussies              : {0}" -f $cOK) -ForegroundColor Green
  Write-Host ("  Echouees              : {0}" -f $cKO) -ForegroundColor $(if ($cKO -gt 0) { "Red" } else { "Green" })
  Write-Host ("  Ignorees (utilisateur): {0}" -f $cSkip) -ForegroundColor $(if ($cSkip -gt 0) { "Yellow" } else { "DarkGray" })
  Write-Host ""
}

Write-Host ("  Duree                 : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  Plan CSV              : {0}" -f $csvFile)
Write-Host ("  LOG                   : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

if ($isDryRun) {
  Write-Host ""
  Write-Host "  Pour executer :" -ForegroundColor DarkGray
  Write-Host "    -Supprimer  : Phase A (types globaux vides)" -ForegroundColor DarkGray
  Write-Host "    -Fusionner  : Phase B (homonymes globaux)" -ForegroundColor DarkGray
  Write-Host "    -Nettoyer   : Phase C (types projet-scoped vides)" -ForegroundColor DarkGray
  Write-Host "    -Rapport    : Phase D (projets Team-managed)" -ForegroundColor DarkGray
  Write-Host "    -Execute    : tout" -ForegroundColor DarkGray
  Write-Host ""
}

Log "Termine."