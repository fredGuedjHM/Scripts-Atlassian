<#
.SYNOPSIS
  Nettoyage-TypesProjetScoped-Jira.ps1
  Supprime les types projet-scoped via conversion Team-managed -> Company-managed.

.DESCRIPTION
  Strategie pour chaque projet concerne :
  1. Convertir le projet Team-managed en Company-managed
     (les types projet-scoped deviennent globaux automatiquement)
  2. Supprimer les types devenus globaux (0 issue)
  3. Log et CSV de toutes les actions

  Ce script cible les projets Team-managed dont les types ont 0 issue
  et qui ne peuvent pas etre supprimes directement via l'API.

  ATTENTION : la conversion Team-managed -> Company-managed est IRREVERSIBLE.
  Pour des projets archives (z_, Z_) avec 0 issue, c'est sans risque.

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+
  Version        : 2 (2026-07-09)

.EXAMPLE
  .\Nettoyage-TypesProjetScoped-Jira.ps1
#>

[CmdletBinding()]
param(
  [int] $MaxRetries = 5
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
$logFile = Join-Path $ExportsDir ("Nettoyage_ProjetScoped_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("Nettoyage_ProjetScoped_{0}.csv" -f $ts)
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

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

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

function Load-SiteCredentials([string]$CredFile, [string]$SiteName) {
  if (-not (Test-Path $CredFile)) {
    Log "Fichier $CredFile introuvable." "ERROR"
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

function Get-JiraErrorMessage([string]$responseBody) {
  if (-not $responseBody) { return "" }
  try {
    $errJson = $responseBody | ConvertFrom-Json
    if ($errJson.errorMessages -and ($errJson.errorMessages | Measure-Object).Count -gt 0) {
      return $errJson.errorMessages -join "; "
    }
    if ($errJson.errors) {
      return ($errJson.errors.PSObject.Properties | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Value }) -join "; "
    }
  } catch {}
  if ($responseBody.Length -gt 200) { return $responseBody.Substring(0, 200) + "..." }
  return $responseBody
}

# ============================================================
# DEBUT
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  NETTOYAGE TYPES PROJET-SCOPED" -ForegroundColor Cyan
Write-Host "  (conversion Company-managed + suppression)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "================================================================"
Log "  NETTOYAGE TYPES PROJET-SCOPED"
Log "  Strategie : conversion Company-managed puis suppression"
Log "================================================================"

$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
Log ("  Site : {0}" -f $site.BaseUrl)

$startTime = Get-Date

# CSV header
$csvCols = @("Etape","Action","ProjectKey","ProjectName","TypeName","TypeId","Result","Detail")
$csvHeaderLine = ($csvCols | ForEach-Object { CsvEscape $_ }) -join ";"
Write-CsvToFile $csvFile $csvHeaderLine -Header

function Write-ActionCsv {
  param([string]$Etape, [string]$Action, [string]$ProjectKey, [string]$ProjectName,
        [string]$TypeName, [string]$TypeId, [string]$Result, [string]$Detail)
  $vals = @($Etape, $Action, $ProjectKey, $ProjectName, $TypeName, $TypeId, $Result, $Detail)
  $line = ($vals | ForEach-Object { CsvEscape $_ }) -join ";"
  Write-CsvToFile $script:csvFile $line
}

# ============================================================
# ETAPE 1 : IDENTIFIER LES PROJETS AVEC TYPES PROJET-SCOPED
# ============================================================

Log "=== ETAPE 1 : Listing des types projet-scoped ==="

$url = "{0}/rest/api/3/issuetype" -f $site.BaseUrl
$resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers
if (-not $resp.ok) { Log ("  Erreur listing : status={0}" -f $resp.status) "ERROR"; return }

$issueTypes = $resp.content | ConvertFrom-Json

# Collecter les types projet-scoped
$projectScopedTypes = New-Object System.Collections.Generic.List[object]
foreach ($it in $issueTypes) {
  if ($it.scope -and $it.scope.type -eq "PROJECT") {
    $projectScopedTypes.Add(@{
      Id             = [string]$it.id
      Name           = [string]$it.name
      IsSubtask      = [bool]$it.subtask
      ScopeProjectId = [string]$it.scope.project.id
    }) | Out-Null
  }
}

# Grouper par projet
$projectTypeGroups = @{}
foreach ($t in $projectScopedTypes) {
  $tProjId = $t.ScopeProjectId
  if (-not $projectTypeGroups.ContainsKey($tProjId)) { $projectTypeGroups[$tProjId] = @() }
  $projectTypeGroups[$tProjId] += $t
}

# Resoudre les noms de projets et identifier le style
$projectInfos = @{}
foreach ($projId in $projectTypeGroups.Keys) {
  $projUrl = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $projId
  $projResp = Invoke-ApiCall -Method "GET" -Url $projUrl -Headers $site.Headers
  if ($projResp.ok) {
    $projJson = $projResp.content | ConvertFrom-Json
    $pStyle = if ($projJson.style) { [string]$projJson.style } else { "classic" }
    $projectInfos[$projId] = @{
      Key   = [string]$projJson.key
      Name  = [string]$projJson.name
      Style = $pStyle
    }
    Log ("  Projet {0} - {1} (style={2})" -f $projJson.key, $projJson.name, $pStyle)
  } else {
    $projectInfos[$projId] = @{ Key="?"; Name="Projet ID $projId"; Style="inconnu" }
    Log ("  Projet ID {0} : inaccessible" -f $projId) "WARN"
  }
  Start-Sleep -Milliseconds 100
}

# Verification du comptage (0 issue par type)
Log "=== Verification du comptage ==="
$projectsToProcess = @{}

foreach ($projId in $projectTypeGroups.Keys) {
  $pTypes = $projectTypeGroups[$projId]
  $allZero = $true
  $typesZero = @()
  $typesNonZero = @()

  foreach ($t in $pTypes) {
    $searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl
    $jqlText = "issuetype = {0}" -f $t.Id
    $bodyObj = @{ jql = $jqlText; maxResults = 1; fields = @("key") }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress
    $searchResp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $bodyJson

    $hasIssues = $false
    if ($searchResp.ok) {
      $searchJson = $searchResp.content | ConvertFrom-Json
      if (($searchJson.issues | Measure-Object).Count -gt 0) { $hasIssues = $true }
    }

    if ($hasIssues) {
      $allZero = $false
      $typesNonZero += $t
    } else {
      $typesZero += $t
    }
  }

  $pInfo = $projectInfos[$projId]
  if ($typesZero.Count -gt 0) {
    $projectsToProcess[$projId] = @{
      TypesToDelete = $typesZero
      TypesWithIssues = $typesNonZero
      AllZero = $allZero
    }
    $projLabel = "{0} ({1})" -f $pInfo.Key, $pInfo.Name
    Log ("  {0} : {1} types a supprimer, {2} types avec issues" -f $projLabel, $typesZero.Count, $typesNonZero.Count)
  }
}

$cProjects = ($projectsToProcess.Keys | Measure-Object).Count
$cTypesTotal = 0
foreach ($projId in $projectsToProcess.Keys) {
  $cTypesTotal += $projectsToProcess[$projId].TypesToDelete.Count
}

if ($cTypesTotal -eq 0) {
  Write-Host "  Aucun type projet-scoped avec 0 issue a supprimer." -ForegroundColor DarkGray
  Log "  Rien a faire."
  return
}

# Affichage
Write-Host ""
Write-Host ("  {0} projets, {1} types a supprimer :" -f $cProjects, $cTypesTotal) -ForegroundColor Magenta
Write-Host ""

foreach ($projId in ($projectsToProcess.Keys | Sort-Object)) {
  $pInfo = $projectInfos[$projId]
  $pData = $projectsToProcess[$projId]
  $projLabel = "{0} ({1})" -f $pInfo.Key, $pInfo.Name
  $cDel = $pData.TypesToDelete.Count
  $cKeep = $pData.TypesWithIssues.Count

  Write-Host ("    Projet {0} [style={1}]" -f $projLabel, $pInfo.Style) -ForegroundColor Magenta
  Write-Host ("      {0} types a supprimer, {1} types avec issues (conserves)" -f $cDel, $cKeep) -ForegroundColor DarkMagenta

  foreach ($t in ($pData.TypesToDelete | Sort-Object { $_.Name })) {
    Write-Host ("        [SUPPRIMER]  {0,-25} (ID:{1})" -f $t.Name, $t.Id) -ForegroundColor Red
  }
  foreach ($t in ($pData.TypesWithIssues | Sort-Object { $_.Name })) {
    Write-Host ("        [CONSERVER]  {0,-25} (ID:{1}) (a des issues)" -f $t.Name, $t.Id) -ForegroundColor DarkGray
  }
  Write-Host ""
}

# Confirmation
$confirmMsg = "Traiter {0} projets ({1} types a supprimer) ?`n`nStrategie pour chaque projet :`n1. Convertir Team-managed -> Company-managed (IRREVERSIBLE)`n2. Relire les types (devenus globaux)`n3. Supprimer les types avec 0 issue`n`nCette operation est irreversible." -f $cProjects, $cTypesTotal
$confirm = [System.Windows.Forms.MessageBox]::Show(
  $confirmMsg,
  "Nettoyage types projet-scoped",
  [System.Windows.Forms.MessageBoxButtons]::YesNo,
  [System.Windows.Forms.MessageBoxIcon]::Warning)

if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
  Write-Host "  Operation annulee." -ForegroundColor Yellow
  Log "  Operation annulee par l'utilisateur."
  return
}

# ============================================================
# ETAPE 2 : TRAITEMENT PAR PROJET
# ============================================================

$totalConverted = 0; $totalDeleted = 0; $totalErrors = 0

foreach ($projId in ($projectsToProcess.Keys | Sort-Object)) {
  $pInfo = $projectInfos[$projId]
  $pData = $projectsToProcess[$projId]
  $projKey  = $pInfo.Key
  $projName = $pInfo.Name
  $projLabel = "{0} ({1})" -f $projKey, $projName

  Write-Host ""
  Write-Host ("  ---- Projet {0} ----" -f $projLabel) -ForegroundColor Magenta
  Log ("=== Traitement projet {0} ===" -f $projLabel)

  # --------------------------------------------------------
  # 2a. Convertir en Company-managed
  # --------------------------------------------------------

  if ($pInfo.Style -eq "next-gen") {
    Log ("  Conversion {0} : Team-managed -> Company-managed" -f $projKey)
    Write-Host ("    Conversion Team-managed -> Company-managed..." -f $projKey) -ForegroundColor DarkYellow

    $convertUrl = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $projKey
    $convertBody = @{
      assigneeType = "UNASSIGNED"
    } | ConvertTo-Json -Depth 5 -Compress

    $convertResp = Invoke-ApiCall -Method "PUT" -Url $convertUrl -Headers $site.Headers -Body $convertBody

    if ($convertResp.ok) {
      $totalConverted++
      Log ("  Projet {0} converti en Company-managed" -f $projKey)
      Write-Host ("    Converti en Company-managed" -f $projKey) -ForegroundColor Green
      Write-ActionCsv "CONVERTIR" "Team->Company" $projKey $projName "" "" "OK" "Converti en Company-managed"
    } else {
      $errMsg = Get-JiraErrorMessage $convertResp.body
      $errDetail = "status={0}, erreur={1}" -f $convertResp.status, $errMsg
      Log ("  ERREUR conversion {0} : {1}" -f $projKey, $errDetail) "ERROR"
      Write-Host ("    ERREUR conversion : {0}" -f $errMsg) -ForegroundColor Red
      Write-ActionCsv "CONVERTIR" "Team->Company" $projKey $projName "" "" "ERREUR" $errDetail

      # Si la conversion echoue, on ne peut pas supprimer les types
      foreach ($t in $pData.TypesToDelete) {
        Write-ActionCsv "SUPPRIMER" "ANNULE" $projKey $projName $t.Name $t.Id "ANNULE" "Conversion du projet echouee"
      }
      $totalErrors += $pData.TypesToDelete.Count
      continue
    }

    # Attendre que Jira propage la conversion
    Start-Sleep -Seconds 2
  } else {
    Log ("  Projet {0} deja en Company-managed (style={1})" -f $projKey, $pInfo.Style)
    Write-Host ("    Deja en Company-managed" -f $projKey) -ForegroundColor DarkGray
  }

  # --------------------------------------------------------
  # 2b. Relire les types pour obtenir les IDs a jour
  # --------------------------------------------------------

  Log ("  Relecture des types apres conversion...")
  $rereadUrl = "{0}/rest/api/3/issuetype" -f $site.BaseUrl
  $rereadResp = Invoke-ApiCall -Method "GET" -Url $rereadUrl -Headers $site.Headers

  $currentTypeIds = @{}
  if ($rereadResp.ok) {
    $rereadTypes = $rereadResp.content | ConvertFrom-Json
    foreach ($rt in $rereadTypes) {
      $currentTypeIds[[string]$rt.id] = @{
        Name    = [string]$rt.name
        Scope   = if ($rt.scope) { [string]$rt.scope.type } else { "GLOBAL" }
        Subtask = [bool]$rt.subtask
      }
    }
  }

  # --------------------------------------------------------
  # 2c. Supprimer les types (maintenant globaux)
  # --------------------------------------------------------

  foreach ($t in ($pData.TypesToDelete | Sort-Object { $_.Name })) {
    # Verifier que le type existe encore
    if (-not $currentTypeIds.ContainsKey($t.Id)) {
      Log ("  Type {0} (ID:{1}) n'existe plus (deja supprime ou fusionne)" -f $t.Name, $t.Id)
      Write-Host ("    DEJA SUPPRIME : {0}" -f $t.Name) -ForegroundColor DarkGray
      Write-ActionCsv "SUPPRIMER" "DEJA_SUPPRIME" $projKey $projName $t.Name $t.Id "OK" "Type n'existe plus apres conversion"
      $totalDeleted++
      continue
    }

    $typeInfo = $currentTypeIds[$t.Id]
    $typeScope = $typeInfo.Scope

    # Verifier que le type est bien devenu global
    if ($typeScope -eq "PROJECT") {
      Log ("  Type {0} (ID:{1}) est encore projet-scoped apres conversion" -f $t.Name, $t.Id) "WARN"
      Write-Host ("    ENCORE PROJET-SCOPED : {0} (tentative de suppression quand meme)" -f $t.Name) -ForegroundColor Yellow
    }

    # Recompter pour securite
    $searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl
    $jqlText = "issuetype = {0}" -f $t.Id
    $bodyObj = @{ jql = $jqlText; maxResults = 1; fields = @("key") }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress
    $searchResp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $bodyJson

    $hasIssues = $false
    if ($searchResp.ok) {
      $searchJson = $searchResp.content | ConvertFrom-Json
      if (($searchJson.issues | Measure-Object).Count -gt 0) { $hasIssues = $true }
    }

    if ($hasIssues) {
      Log ("  Type {0} (ID:{1}) a des issues apres conversion, suppression annulee" -f $t.Name, $t.Id) "WARN"
      Write-Host ("    IGNORE (a des issues) : {0}" -f $t.Name) -ForegroundColor Yellow
      Write-ActionCsv "SUPPRIMER" "ANNULE" $projKey $projName $t.Name $t.Id "ANNULE" "Issues trouvees apres conversion"
      $totalErrors++
      continue
    }

    # Suppression
    $delUrl = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $t.Id
    $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers

    if ($delResp.ok -or $delResp.status -eq 204) {
      $totalDeleted++
      Log ("  SUPPRIME : {0} (ID:{1}) [{2}]" -f $t.Name, $t.Id, $projLabel)
      Write-Host ("    SUPPRIME : {0}" -f $t.Name) -ForegroundColor Green
      Write-ActionCsv "SUPPRIMER" "DELETE" $projKey $projName $t.Name $t.Id "OK" "Supprime apres conversion"
    } else {
      $totalErrors++
      $errMsg = Get-JiraErrorMessage $delResp.body
      $errDetail = "status={0}, erreur={1}" -f $delResp.status, $errMsg
      Log ("  ERREUR suppression {0} : {1}" -f $t.Name, $errDetail) "ERROR"
      Write-Host ("    ERREUR : {0} : {1}" -f $t.Name, $errMsg) -ForegroundColor Red
      Write-ActionCsv "SUPPRIMER" "DELETE" $projKey $projName $t.Name $t.Id "ERREUR" $errDetail
    }
    Start-Sleep -Milliseconds 200
  }
}

# ============================================================
# RESUME
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  NETTOYAGE TERMINE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Projets convertis     : {0}" -f $totalConverted)
Write-Host ("  Types supprimes       : {0}" -f $totalDeleted)
Write-Host ("  Erreurs               : {0}" -f $totalErrors)
Write-Host ("  Duree                 : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  CSV                   : {0}" -f $csvFile)
Write-Host ("  LOG                   : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "Nettoyage types projet-scoped termine."