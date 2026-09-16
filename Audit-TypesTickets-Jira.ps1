<#
.SYNOPSIS
  Audit-TypesTickets-Jira.ps1
  Audite les types de tickets Jira : homonymes, inutilises, projet-scoped vides, Team-managed.

.DESCRIPTION
  Ce script :
  1. Liste tous les types de tickets
  2. Resout les noms des projets (libelles)
  3. Compte l'utilisation de chaque type
  4. Detecte les homonymes globaux
  5. Identifie les types avec 0 issue (globaux et projet-scoped)
  6. Recense les projets Team-managed avec types isoles
  7. Exporte un plan d'action CSV

  Aucune modification n'est effectuee. Utilisez Nettoyage-TypesTickets-Jira.ps1 pour agir.

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+
  Version        : 4 (2026-07-09)

.EXAMPLE
  .\Audit-TypesTickets-Jira.ps1
#>

[CmdletBinding()]
param(
  [int] $MaxRetries = 5,
  [int] $MaxCount   = 10000
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("Audit_Types_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("Audit_Types_{0}.csv" -f $ts)
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
    Log "Fichier $CredFile introuvable." "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
      ("Fichier de credentials introuvable :`n{0}`n`nCreez-le avec Export-Clixml ou copiez-le depuis un autre script." -f $CredFile),
      "Credentials manquants",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    throw "Credentials introuvables : $CredFile"
  }
  $data  = Import-Clixml -Path $CredFile
  $url   = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  Log ("  Credentials charges depuis {0}" -f $CredFile)
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

# Helper : somme des IssueCount sur une collection de hashtables
function Sum-IssueCount($items) {
  $total = 0
  foreach ($item in $items) { $total += $item.IssueCount }
  return $total
}

# ============================================================
# DEBUT
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  AUDIT DES TYPES DE TICKETS JIRA" -ForegroundColor Cyan
Write-Host "  Mode : ANALYSE (aucune modification)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "================================================================"
Log "  AUDIT DES TYPES DE TICKETS JIRA"
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
# ETAPE 1b : RESOLUTION DES NOMS DE PROJETS
# ============================================================

Log "=== ETAPE 1b : Resolution des noms de projets ==="

$projectTypes = $typeData | Where-Object { $_.Scope -eq "PROJECT" }
$uniqueProjectIds = $projectTypes | ForEach-Object { $_.ScopeProjectId } | Sort-Object -Unique

$projectInfos = @{}
$cProj = ($uniqueProjectIds | Measure-Object).Count
$cProjDone = 0

foreach ($projId in $uniqueProjectIds) {
  $cProjDone++
  $projUrl = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $projId
  $projResp = Invoke-ApiCall -Method "GET" -Url $projUrl -Headers $site.Headers
  if ($projResp.ok) {
    $projJson = $projResp.content | ConvertFrom-Json
    $pStyle = if ($projJson.style) { [string]$projJson.style } else { "classic" }
    $pLead  = if ($projJson.lead -and $projJson.lead.displayName) { [string]$projJson.lead.displayName } else { "" }
    $projectInfos[$projId] = @{
      Key   = [string]$projJson.key
      Name  = [string]$projJson.name
      Style = $pStyle
      Lead  = $pLead
    }
    Log ("  {0}/{1} Projet {2} - {3} (style={4}, lead={5})" -f $cProjDone, $cProj, $projJson.key, $projJson.name, $pStyle, $pLead)
  } else {
    $projectInfos[$projId] = @{ Key="?"; Name="Projet ID $projId"; Style="inconnu"; Lead="" }
    Log ("  {0}/{1} Projet ID {2} : inaccessible (status={3})" -f $cProjDone, $cProj, $projId, $projResp.status) "WARN"
  }
  Start-Sleep -Milliseconds 100
}

Log ("  {0} projets resolus" -f $cProj)

function Get-ProjectLabel([string]$projectId) {
  if ($script:projectInfos.ContainsKey($projectId)) {
    $pi = $script:projectInfos[$projectId]
    return "{0} ({1})" -f $pi.Key, $pi.Name
  }
  return "ID:$projectId"
}

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

  if ($t.Scope -eq "PROJECT") {
    $projLabel = Get-ProjectLabel $t.ScopeProjectId
    $scopeLabel = " [Projet: $projLabel]"
  } else {
    $scopeLabel = " [Global]"
  }
  $countLabel = if ($count -ge $MaxCount) { "{0}+" -f $count } else { "$count" }
  Log ("  {0,-35} (ID:{1}){2} : {3}" -f $t.Name, $t.Id, $scopeLabel, $countLabel)
}
Write-Progress -Activity "Comptage" -Completed

# ============================================================
# ETAPE 3 : GROUPES D'HOMONYMES
# ============================================================

Log "=== ETAPE 3 : Groupes d'homonymes ==="

$globalTypes = $typeData | Where-Object { $_.Scope -ne "PROJECT" }

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
# ETAPE 3b : GROUPEMENT DES TYPES PAR PROJET
# ============================================================

Log "=== ETAPE 3b : Groupement par projet ==="

$projectTypeGroups = @{}
foreach ($t in $projectTypes) {
  $tProjId = $t.ScopeProjectId
  if (-not $projectTypeGroups.ContainsKey($tProjId)) {
    $projectTypeGroups[$tProjId] = @()
  }
  $projectTypeGroups[$tProjId] += $t
}

$cTeamManaged = ($projectTypeGroups.Keys | Measure-Object).Count
Log ("  {0} projets avec types projet-scoped" -f $cTeamManaged)

# ============================================================
# ETAPE 4 : PLAN D'ACTION
# ============================================================

Log "=== ETAPE 4 : Plan d'action ==="

$plan = New-Object System.Collections.Generic.List[object]

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
      Phase       = "A-SUPPRIMER"
      Action      = "SUPPRIMER"
      SourceName  = $t.Name
      SourceId    = $t.Id
      SourceScope = "Global"
      SourceCount = 0
      TargetName  = ""
      TargetId    = ""
      ProjectKey  = ""
      ProjectName = ""
      Reason      = "0 issue, type global inutilise"
      Status      = "PLANIFIE"
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
        Phase       = "B-FUSIONNER"
        Action      = "SUPPRIMER"
        SourceName  = $src.Name
        SourceId    = $src.Id
        SourceScope = "Global"
        SourceCount = 0
        TargetName  = $target.Name
        TargetId    = $target.Id
        ProjectKey  = ""
        ProjectName = ""
        Reason      = $reasonText
        Status      = "PLANIFIE"
      }) | Out-Null
    } else {
      $reasonText = "Homonyme de '{0}' (ID:{1}, {2} issues), migrer {3} issues" -f $target.Name, $target.Id, $target.IssueCount, $src.IssueCount
      $plan.Add(@{
        Phase       = "B-FUSIONNER"
        Action      = "MIGRER+SUPPRIMER"
        SourceName  = $src.Name
        SourceId    = $src.Id
        SourceScope = "Global"
        SourceCount = $src.IssueCount
        TargetName  = $target.Name
        TargetId    = $target.Id
        ProjectKey  = ""
        ProjectName = ""
        Reason      = $reasonText
        Status      = "PLANIFIE"
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
    $reasonText = "Type projet-scoped ({0} - {1}), 0 issue" -f $projKey, $projName
    $scopeText  = "project:{0}" -f $t.ScopeProjectId
    $plan.Add(@{
      Phase       = "C-NETTOYER"
      Action      = "SUPPRIMER"
      SourceName  = $t.Name
      SourceId    = $t.Id
      SourceScope = $scopeText
      SourceCount = 0
      TargetName  = ""
      TargetId    = ""
      ProjectKey  = $projKey
      ProjectName = $projName
      Reason      = $reasonText
      Status      = "PLANIFIE"
    }) | Out-Null
  }
}

# --- Phase D : Rapport projets Team-managed ---
foreach ($projId in ($projectTypeGroups.Keys | Sort-Object)) {
  $pTypes = $projectTypeGroups[$projId]
  $pInfo  = $projectInfos[$projId]
  $projKey   = if ($pInfo) { $pInfo.Key } else { "?" }
  $projName  = if ($pInfo) { $pInfo.Name } else { "?" }
  $projStyle = if ($pInfo) { $pInfo.Style } else { "?" }
  $projLead  = if ($pInfo) { $pInfo.Lead } else { "?" }

  $cTypesProj  = ($pTypes | Measure-Object).Count
  $cTypesVides = ($pTypes | Where-Object { $_.IssueCount -eq 0 } | Measure-Object).Count
  $cTypesActifs= $cTypesProj - $cTypesVides
  $typeNames   = ($pTypes | ForEach-Object { $_.Name } | Sort-Object -Unique) -join ", "
  $totalIssues = Sum-IssueCount $pTypes

  $actionText = if ($cTypesActifs -eq 0) { "CONVERTIR (aucune issue)" } else { "CONVERTIR (migration requise)" }

  $reasonText = "Projet {0} ({1}), style={2}, lead={3}, {4} types ({5} vides, {6} actifs), {7} issues, types: {8}" -f `
    $projKey, $projName, $projStyle, $projLead, $cTypesProj, $cTypesVides, $cTypesActifs, $totalIssues, $typeNames

  $plan.Add(@{
    Phase       = "D-RAPPORT"
    Action      = $actionText
    SourceName  = $projKey
    SourceId    = $projId
    SourceScope = "Team-managed"
    SourceCount = $totalIssues
    TargetName  = ""
    TargetId    = ""
    ProjectKey  = $projKey
    ProjectName = $projName
    Reason      = $reasonText
    Status      = "A EVALUER"
  }) | Out-Null
}

$cPhaseA    = ($plan | Where-Object { $_.Phase -eq "A-SUPPRIMER" } | Measure-Object).Count
$cPhaseBSup = ($plan | Where-Object { $_.Phase -eq "B-FUSIONNER" -and $_.Action -eq "SUPPRIMER" } | Measure-Object).Count
$cPhaseBMig = ($plan | Where-Object { $_.Phase -eq "B-FUSIONNER" -and $_.Action -eq "MIGRER+SUPPRIMER" } | Measure-Object).Count
$cPhaseC    = ($plan | Where-Object { $_.Phase -eq "C-NETTOYER" } | Measure-Object).Count
$cPhaseD    = ($plan | Where-Object { $_.Phase -eq "D-RAPPORT" } | Measure-Object).Count

# ============================================================
# ETAPE 5 : EXPORT CSV
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
Write-Host "  PLAN D'ACTION (AUDIT)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Phase A ---
Write-Host ""
Write-Host ("  PHASE A - TYPES GLOBAUX INUTILISES ({0}) :" -f $cPhaseA) -ForegroundColor Red
if ($cPhaseA -eq 0) {
  Write-Host "    Aucun." -ForegroundColor DarkGray
} else {
  foreach ($p in ($plan | Where-Object { $_.Phase -eq "A-SUPPRIMER" } | Sort-Object { $_.SourceName })) {
    Write-Host ("    [SUPPRIMER]  {0,-35} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor Red
  }
}

# --- Phase B ---
$cPhaseB = $cPhaseBSup + $cPhaseBMig
Write-Host ""
Write-Host ("  PHASE B - HOMONYMES GLOBAUX ({0} actions) :" -f $cPhaseB) -ForegroundColor DarkYellow
if ($cPhaseB -eq 0) {
  Write-Host "    Aucun." -ForegroundColor DarkGray
} else {
  foreach ($p in ($plan | Where-Object { $_.Phase -eq "B-FUSIONNER" } | Sort-Object { $_.SourceName })) {
    if ($p.Action -eq "MIGRER+SUPPRIMER") {
      Write-Host ("    [MIGRER]     {0,-30} (ID:{1}, {2} issues) -> {3} (ID:{4})" -f $p.SourceName, $p.SourceId, $p.SourceCount, $p.TargetName, $p.TargetId) -ForegroundColor DarkYellow
    } else {
      Write-Host ("    [SUPPRIMER]  {0,-30} (ID:{1}, 0 issue) [homonyme de {2}]" -f $p.SourceName, $p.SourceId, $p.TargetName) -ForegroundColor Red
    }
  }
}

# --- Phase C ---
Write-Host ""
Write-Host ("  PHASE C - TYPES PROJET-SCOPED VIDES ({0}) :" -f $cPhaseC) -ForegroundColor Magenta
if ($cPhaseC -eq 0) {
  Write-Host "    Aucun." -ForegroundColor DarkGray
} else {
  $phaseCByProject = $plan | Where-Object { $_.Phase -eq "C-NETTOYER" } | Group-Object { $_.ProjectKey } | Sort-Object Name
  foreach ($grp in $phaseCByProject) {
    $projKey  = $grp.Name
    $projName = $grp.Group[0].ProjectName
    $grpCount = ($grp.Group | Measure-Object).Count
    Write-Host ("    Projet {0} ({1}) - {2} types :" -f $projKey, $projName, $grpCount) -ForegroundColor Magenta
    foreach ($p in ($grp.Group | Sort-Object { $_.SourceName })) {
      Write-Host ("      [SUPPRIMER]  {0,-30} (ID:{1})" -f $p.SourceName, $p.SourceId) -ForegroundColor DarkMagenta
    }
  }
}

# --- Phase D ---
Write-Host ""
Write-Host ("  PHASE D - PROJETS TEAM-MANAGED ({0}) :" -f $cPhaseD) -ForegroundColor Blue
if ($cPhaseD -eq 0) {
  Write-Host "    Aucun." -ForegroundColor DarkGray
} else {
  foreach ($p in ($plan | Where-Object { $_.Phase -eq "D-RAPPORT" } | Sort-Object { $_.SourceName })) {
    $pTypes = $projectTypeGroups[$p.SourceId]
    $cTypesProj  = ($pTypes | Measure-Object).Count
    $cTypesVides = ($pTypes | Where-Object { $_.IssueCount -eq 0 } | Measure-Object).Count
    $pInfo = $projectInfos[$p.SourceId]
    $projLead = if ($pInfo) { $pInfo.Lead } else { "?" }
    $color = if ($p.SourceCount -eq 0) { "Green" } else { "Yellow" }
    $riskLabel = if ($p.SourceCount -eq 0) { "SANS RISQUE" } else { "MIGRATION REQUISE" }

    Write-Host ("    {0,-10} {1,-30} lead={2}" -f $p.SourceName, $p.ProjectName, $projLead) -ForegroundColor $color
    Write-Host ("               {0} issues, {1} types ({2} vides) -> {3}" -f $p.SourceCount, $cTypesProj, $cTypesVides, $riskLabel) -ForegroundColor $color

    foreach ($pt in ($pTypes | Sort-Object { $_.Name })) {
      $ptCount = if ($pt.IssueCount -ge $MaxCount) { "{0}+" -f $pt.IssueCount } else { [string]$pt.IssueCount }
      $ptColor = if ($pt.IssueCount -eq 0) { "DarkGray" } else { "White" }
      Write-Host ("                 - {0,-25} (ID:{1}) {2} issues" -f $pt.Name, $pt.Id, $ptCount) -ForegroundColor $ptColor
    }
    Write-Host ""
  }
}

# ============================================================
# RESUME
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME DE L'AUDIT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Types analyses        : {0}" -f $cTypes)
Write-Host ("    - Globaux           : {0}" -f ($globalTypes | Measure-Object).Count)
Write-Host ("    - Projet-scoped     : {0}" -f ($projectTypes | Measure-Object).Count)
Write-Host ""
Write-Host ("  Phase A planifiees    : {0} suppressions globales" -f $cPhaseA) -ForegroundColor Red
Write-Host ("  Phase B planifiees    : {0} fusions ({1} suppr + {2} migr)" -f ($cPhaseBSup + $cPhaseBMig), $cPhaseBSup, $cPhaseBMig) -ForegroundColor DarkYellow
Write-Host ("  Phase C planifiees    : {0} suppressions projet-scoped" -f $cPhaseC) -ForegroundColor Magenta
Write-Host ("  Phase D               : {0} projets Team-managed" -f $cPhaseD) -ForegroundColor Blue
Write-Host ""
Write-Host ("  Duree                 : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  Plan CSV              : {0}" -f $csvFile)
Write-Host ("  LOG                   : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Utilisez Nettoyage-TypesTickets-Jira.ps1 pour executer les actions." -ForegroundColor DarkGray
Write-Host ""

Log "Audit termine."