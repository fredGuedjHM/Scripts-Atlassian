<#
.SYNOPSIS
  Gestion-Datasources-Analytics.ps1
  Inventaire des projets Jira par categorie, generation des requetes SQL
  pour Atlassian Analytics Data Lake, et maintenance de la coherence.

.DESCRIPTION
  Ce script :
  1. Liste tous les projets Jira avec leurs categories
  2. Detecte les anomalies (projets sans categorie, categories vides)
  3. Regroupe les projets en datasources logiques :
     - 1 datasource TRANSVERSE (categories globales : DSIM, GOUVERNANCE, HM, DSIT)
     - 1 datasource par categorie metier (DSIM_CED, DSIM_VDC, etc.)
  4. Genere les requetes SQL pretes a coller dans Atlassian Analytics
  5. Mode correction : reassigner les categories de projets

  MODES :
    -Mode Audit       => Inventaire + coherence + generation requetes (defaut)
    -Mode Correction   => Audit + reassignation interactive des projets orphelins

  CREDENTIALS :
    secrets\site-admin.xml : SiteUrl + Email + API Token

  FICHIERS GENERES (dans exports\) :
    Datasources_Analytics_{ts}.csv   => Inventaire projets/categories/datasource
    Datasources_Analytics_{ts}.sql   => Requetes SQL pour Atlassian Analytics
    Datasources_Analytics_{ts}.log   => Journal d'execution
    Datasources_Coherence_{ts}.csv   => Rapport d'anomalies

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Gestion-Datasources-Analytics.ps1
  .\Gestion-Datasources-Analytics.ps1 -Mode Correction
  .\Gestion-Datasources-Analytics.ps1 -Mode Audit -IncludeArchived $false
#>

[CmdletBinding()]
param(
  [ValidateSet("Audit","Correction")]
  [string] $Mode = "Audit",

  [bool] $IncludeArchived = $true,

  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ============================================================
# CONFIGURATION DES DATASOURCES
# ============================================================
# Categories qui alimentent la datasource TRANSVERSE (globale)
# Modifiez cette liste selon votre organisation
$transverseCategories = @(
  "DSIM",
  "DSIT",
  "GOUVERNANCE",
  "HM"
)

# Categories a exclure des datasources metier (archivage)
$excludedCategories = @(
  "z-Projets archivés",
  "z-Projets archiv\u00e9s",
  "A archiver (Paramétrage)",
  "A archiver (Param\u00e9trage)",
  "A maintenir (en att; finalisation migration)"
)

# Colonnes SQL a inclure dans les requetes Analytics
$sqlColumns = @(
  "i.key AS issue_key",
  "i.summary",
  "p.key AS project_key",
  "p.name AS project_name",
  "i.issue_type_name",
  "i.status_name",
  "i.status_category_name",
  "i.priority_name",
  "i.assignee_display_name",
  "i.reporter_display_name",
  "i.created",
  "i.updated",
  "i.resolved",
  "i.due_date",
  "i.story_points",
  "i.sprint_name",
  "i.labels"
)

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile       = Join-Path $ExportsDir ("Datasources_Analytics_{0}.log" -f $ts)
$csvFile       = Join-Path $ExportsDir ("Datasources_Analytics_{0}.csv" -f $ts)
$sqlFile       = Join-Path $ExportsDir ("Datasources_Analytics_{0}.sql" -f $ts)
$coherenceFile = Join-Path $ExportsDir ("Datasources_Coherence_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Write-FileUtf8([string]$path, [string]$content, [bool]$append=$false) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  if ($append) {
    [System.IO.File]::AppendAllText($path, "$content`r`n", $utf8Bom)
  } else {
    [System.IO.File]::WriteAllText($path, "$content`r`n", $utf8Bom)
  }
}

# -------------------- Network --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# -------------------- Credentials --------------------
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
  $adminEmail = Read-Host "Email administrateur"
  $apiTokenSecure = Read-Host "API Token" -AsSecureString
  @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $siteCredFile
  Log "Credentials sauvegardes dans $siteCredFile"
}

$siteData = Import-Clixml -Path $siteCredFile
$siteUrl = [string]$siteData.SiteUrl
$siteEmail = [string]$siteData.Email
$siteToken = [System.Net.NetworkCredential]::new("", $siteData.ApiTokenSecureString).Password
$base64Auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${siteEmail}:${siteToken}"))
$jiraHeaders = @{ Authorization = "Basic $base64Auth"; Accept = "application/json" }
$baseUrl = "https://$siteUrl"

# -------------------- HTTP helper --------------------
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
# ETAPE 1 : INVENTAIRE DES CATEGORIES
# ============================================================

Log "================================================================"
Log "  GESTION DATASOURCES ATLASSIAN ANALYTICS"
Log "  Mode: $Mode | Site: $baseUrl"
Log "================================================================"

Log "=== ETAPE 1 : Recuperation des categories de projets ==="
$url = "${baseUrl}/rest/api/3/projectCategory"
$resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $jiraHeaders
if (-not $resp.ok) { throw ("Erreur API categories: status={0}" -f $resp.status) }
$allCategories = $resp.content | ConvertFrom-Json

# Normaliser les noms (gerer les encodages multiples)
$categoryMap = @{} # id => name (normalise)
foreach ($cat in $allCategories) {
  $categoryMap[[string]$cat.id] = [string]$cat.name
}
Log ("  {0} categories trouvees" -f $allCategories.Count)
foreach ($cat in $allCategories) {
  Log ("    ID={0} | {1}" -f $cat.id, $cat.name)
}

# ============================================================
# ETAPE 2 : INVENTAIRE COMPLET DES PROJETS
# ============================================================

Log "=== ETAPE 2 : Recuperation de tous les projets ==="
$allProjects = New-Object System.Collections.Generic.List[object]
$startAt = 0; $pageSize = 50

while ($true) {
  $url = "${baseUrl}/rest/api/3/project/search?maxResults=${pageSize}&startAt=${startAt}&expand=projectCategory,lead"
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $jiraHeaders
  if (-not $resp.ok) { Log ("Erreur API projets start={0} : status={1}" -f $startAt, $resp.status) "ERROR"; break }

  $json = $resp.content | ConvertFrom-Json
  foreach ($p in $json.values) {
    $catName = ""
    $catId = ""
    if ($p.projectCategory) {
      $catName = [string]$p.projectCategory.name
      $catId = [string]$p.projectCategory.id
    }

    $leadName = ""
    if ($p.lead) { $leadName = [string]$p.lead.displayName }

    $pType = "software"
    if ($p.projectTypeKey) { $pType = [string]$p.projectTypeKey }

    $allProjects.Add(@{
      key         = [string]$p.key
      name        = [string]$p.name
      categoryId  = $catId
      category    = $catName
      lead        = $leadName
      type        = $pType
      style       = if ($p.style) { [string]$p.style } else { "" }
      archived    = if ($p.archived -eq $true) { $true } else { $false }
    }) | Out-Null
  }

  Log ("  Projets charges : {0}/{1}" -f $allProjects.Count, $json.total)
  if ($allProjects.Count -ge $json.total) { break }
  $startAt += $pageSize; Start-Sleep -Milliseconds 200
}

Log ("Total projets : {0}" -f $allProjects.Count)

# ============================================================
# ETAPE 3 : CLASSIFICATION EN DATASOURCES
# ============================================================

Log "=== ETAPE 3 : Classification des projets en datasources ==="

# Determiner la datasource de chaque projet
foreach ($proj in $allProjects) {
  $cat = $proj.category

  if (-not $cat) {
    $proj["datasource"] = "__ORPHELIN__"
    $proj["dsType"] = "Orphelin"
  }
  elseif ($excludedCategories | Where-Object { $cat -like $_ -or $cat -match [regex]::Escape($_) }) {
    $proj["datasource"] = "__ARCHIVE__"
    $proj["dsType"] = "Archive"
  }
  elseif ($transverseCategories -contains $cat) {
    $proj["datasource"] = "TRANSVERSE"
    $proj["dsType"] = "Transverse"
  }
  else {
    $proj["datasource"] = $cat
    $proj["dsType"] = "Metier"
  }
}

# Regrouper par datasource
$datasources = @{}
foreach ($proj in $allProjects) {
  $ds = $proj.datasource
  if (-not $datasources.ContainsKey($ds)) {
    $datasources[$ds] = New-Object System.Collections.Generic.List[object]
  }
  $datasources[$ds].Add($proj)
}

# Afficher le resume
Log "--- Datasources identifiees ---"
$dsTransverse = @(); $dsMetier = @(); $dsArchive = @(); $dsOrphelin = @()

foreach ($dsName in ($datasources.Keys | Sort-Object)) {
  $projets = $datasources[$dsName]
  $keys = ($projets | ForEach-Object { $_.key }) -join ", "
  $count = $projets.Count

  switch ($dsName) {
    "__ORPHELIN__" {
      Log ("  [ORPHELIN]    {0} projets : {1}" -f $count, $keys) "WARN"
      $dsOrphelin = $projets
    }
    "__ARCHIVE__" {
      Log ("  [ARCHIVE]     {0} projets : {1}" -f $count, $keys)
      $dsArchive = $projets
    }
    "TRANSVERSE" {
      Log ("  [TRANSVERSE]  {0} projets : {1}" -f $count, $keys)
      $dsTransverse = $projets
    }
    default {
      Log ("  [METIER]      {0} ({1} projets) : {2}" -f $dsName, $count, $keys)
      $dsMetier += @{ name=$dsName; count=$count; keys=$keys; projets=$projets }
    }
  }
}

# ============================================================
# ETAPE 4 : RAPPORT DE COHERENCE
# ============================================================

Log "=== ETAPE 4 : Analyse de coherence ==="

$anomalies = New-Object System.Collections.Generic.List[object]

# 4a. Projets sans categorie
foreach ($proj in $allProjects) {
  if (-not $proj.category) {
    $anomalies.Add(@{
      Type="Projet sans categorie"; Projet=$proj.key; Detail=$proj.name; Suggestion="Assigner une categorie"
    }) | Out-Null
  }
}

# 4b. Categories vides (aucun projet actif)
$usedCategories = $allProjects | Where-Object { $_.category } | ForEach-Object { $_.category } | Sort-Object -Unique
foreach ($cat in $allCategories) {
  $catName = [string]$cat.name
  if ($catName -notin $usedCategories) {
    $anomalies.Add(@{
      Type="Categorie vide"; Projet=""; Detail=$catName; Suggestion="Supprimer ou archiver la categorie"
    }) | Out-Null
  }
}

# 4c. Projets archives encore dans une categorie active
foreach ($proj in $allProjects) {
  if ($proj.archived -and $proj.dsType -notin @("Archive")) {
    $anomalies.Add(@{
      Type="Projet archive hors categorie archive"; Projet=$proj.key
      Detail=("Categorie actuelle: {0}" -f $proj.category)
      Suggestion="Deplacer vers 'z-Projets archives'"
    }) | Out-Null
  }
}

# 4d. Projets dans categorie archive mais pas marques archived
foreach ($proj in $allProjects) {
  if ($proj.dsType -eq "Archive" -and -not $proj.archived) {
    $anomalies.Add(@{
      Type="Projet en categorie archive mais actif"; Projet=$proj.key
      Detail=("Categorie: {0}" -f $proj.category)
      Suggestion="Archiver le projet ou le reclasser"
    }) | Out-Null
  }
}

# 4e. Datasources metier avec un seul projet (risque d'isolation)
foreach ($ds in $dsMetier) {
  if ($ds.count -eq 1) {
    $anomalies.Add(@{
      Type="Datasource avec 1 seul projet"; Projet=$ds.keys
      Detail=("Datasource: {0}" -f $ds.name)
      Suggestion="Verifier si ce projet ne devrait pas etre dans une datasource plus large"
    }) | Out-Null
  }
}

Log ("  {0} anomalies detectees" -f $anomalies.Count)
foreach ($a in $anomalies) {
  Log ("    [{0}] {1} - {2}" -f $a.Type, $a.Projet, $a.Detail) "WARN"
}

# Exporter le rapport de coherence
if ($anomalies.Count -gt 0) {
  $cohColumns = @("Type","Projet","Detail","Suggestion")
  $cohHeader = ($cohColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
  Write-FileUtf8 -path $coherenceFile -content $cohHeader
  foreach ($a in $anomalies) {
    $line = ($cohColumns | ForEach-Object { '"{0}"' -f ([string]$a[$_] -replace '"','""') }) -join ";"
    Write-FileUtf8 -path $coherenceFile -content $line -append $true
  }
  Log "  Rapport de coherence -> $coherenceFile"
}

# ============================================================
# ETAPE 5 : EXPORT CSV INVENTAIRE
# ============================================================

Log "=== ETAPE 5 : Export CSV inventaire ==="

$csvColumns = @("ProjectKey","ProjectName","ProjectType","Lead","Category","Datasource","DatasourceType","Archived")
$csvHeader = ($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
Write-FileUtf8 -path $csvFile -content $csvHeader

foreach ($proj in ($allProjects | Sort-Object { $_.datasource }, { $_.key })) {
  $row = [ordered]@{
    ProjectKey=$proj.key; ProjectName=$proj.name; ProjectType=$proj.type; Lead=$proj.lead
    Category=$proj.category; Datasource=$proj.datasource; DatasourceType=$proj.dsType
    Archived=if ($proj.archived) { "Oui" } else { "Non" }
  }
  $line = ($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
  Write-FileUtf8 -path $csvFile -content $line -append $true
}

Log "  Inventaire CSV -> $csvFile"

# ============================================================
# ETAPE 6 : GENERATION DES REQUETES SQL ANALYTICS
# ============================================================

Log "=== ETAPE 6 : Generation des requetes SQL pour Atlassian Analytics ==="

$sqlColumnsStr = $sqlColumns -join ",`n    "
$sqlContent = New-Object System.Text.StringBuilder

# En-tete du fichier SQL
[void]$sqlContent.AppendLine("-- ================================================================")
[void]$sqlContent.AppendLine("-- REQUETES SQL POUR ATLASSIAN ANALYTICS - DATA LAKE")
[void]$sqlContent.AppendLine(("-- Genere le {0}" -f (Get-Date).ToString("dd/MM/yyyy HH:mm")))
[void]$sqlContent.AppendLine(("-- Site : {0}" -f $baseUrl))
[void]$sqlContent.AppendLine(("-- Projets : {0} | Datasources : {1}" -f $allProjects.Count, ($datasources.Keys | Where-Object { $_ -notin @("__ORPHELIN__","__ARCHIVE__") }).Count))
[void]$sqlContent.AppendLine("-- ================================================================")
[void]$sqlContent.AppendLine("-- INSTRUCTIONS :")
[void]$sqlContent.AppendLine("--   1. Ouvrez Atlassian Analytics > Data Sources")
[void]$sqlContent.AppendLine("--   2. Cliquez 'Create data source' > 'SQL query'")
[void]$sqlContent.AppendLine("--   3. Copiez-collez la requete correspondante")
[void]$sqlContent.AppendLine("--   4. Nommez la datasource selon le commentaire au-dessus")
[void]$sqlContent.AppendLine("--   5. Enregistrez et testez")
[void]$sqlContent.AppendLine("-- ================================================================")
[void]$sqlContent.AppendLine("")

# Fonction pour generer une requete SQL
function New-SqlQuery {
  param([string]$DatasourceName, [string]$Description, [string[]]$ProjectKeys, [string]$FilterType)

  $keysQuoted = ($ProjectKeys | Sort-Object | ForEach-Object { "'$_'" }) -join ", "

  $sql = @"
-- ================================================================
-- DATASOURCE : $DatasourceName
-- Description : $Description
-- Projets ($($ProjectKeys.Count)) : $($ProjectKeys -join ', ')
-- Filtre : $FilterType
-- ================================================================

SELECT
    $sqlColumnsStr
FROM jira_issues i
JOIN jira_projects p ON i.project_id = p.id
WHERE p.key IN ($keysQuoted)
ORDER BY i.updated DESC;

"@
  return $sql
}

# 6a. Datasource TRANSVERSE
if ($dsTransverse.Count -gt 0) {
  $transKeys = $dsTransverse | ForEach-Object { $_.key }
  $transCats = ($dsTransverse | ForEach-Object { $_.category } | Sort-Object -Unique) -join ", "
  $sql = New-SqlQuery `
    -DatasourceName "DS_TRANSVERSE" `
    -Description "Projets transverses et administration ($transCats)" `
    -ProjectKeys $transKeys `
    -FilterType "Categories transverses"
  [void]$sqlContent.Append($sql)
  Log ("  DS_TRANSVERSE : {0} projets ({1})" -f $transKeys.Count, $transCats)
}

# 6b. Datasources METIER (une par categorie)
foreach ($ds in ($dsMetier | Sort-Object { $_.name })) {
  $metierKeys = $ds.projets | ForEach-Object { $_.key }
  $dsCleanName = "DS_" + ($ds.name -replace "[^A-Za-z0-9_]", "_").ToUpper()
  $sql = New-SqlQuery `
    -DatasourceName $dsCleanName `
    -Description ("Projets metier categorie {0}" -f $ds.name) `
    -ProjectKeys $metierKeys `
    -FilterType ("Categorie = {0}" -f $ds.name)
  [void]$sqlContent.Append($sql)
  Log ("  {0} : {1} projets" -f $dsCleanName, $metierKeys.Count)
}

# 6c. Datasource ARCHIVE (optionnelle)
if ($IncludeArchived -and $dsArchive.Count -gt 0) {
  $archKeys = $dsArchive | ForEach-Object { $_.key }
  $sql = New-SqlQuery `
    -DatasourceName "DS_ARCHIVES" `
    -Description "Projets archives (consultation historique)" `
    -ProjectKeys $archKeys `
    -FilterType "Categories d'archivage"
  [void]$sqlContent.Append($sql)
  Log ("  DS_ARCHIVES : {0} projets" -f $archKeys.Count)
}

# 6d. Datasource GLOBALE (tous les projets actifs)
$activeProjects = $allProjects | Where-Object { $_.dsType -in @("Transverse","Metier") }
if ($activeProjects.Count -gt 0) {
  $allActiveKeys = $activeProjects | ForEach-Object { $_.key }
  $sql = New-SqlQuery `
    -DatasourceName "DS_GLOBALE_TOUS_PROJETS" `
    -Description "Tous les projets actifs (toutes categories hors archives)" `
    -ProjectKeys $allActiveKeys `
    -FilterType "Tous projets actifs"
  [void]$sqlContent.Append($sql)
  Log ("  DS_GLOBALE : {0} projets actifs" -f $allActiveKeys.Count)
}

# 6e. Requete de controle : repartition par categorie
$controlSql = @"
-- ================================================================
-- REQUETE DE CONTROLE : Repartition des issues par categorie
-- Utilisez cette requete pour verifier la volumetrie de chaque datasource
-- ================================================================

SELECT
    p.project_category_name AS categorie,
    COUNT(DISTINCT p.key) AS nb_projets,
    COUNT(i.key) AS nb_issues,
    MIN(i.created) AS premiere_issue,
    MAX(i.updated) AS derniere_activite
FROM jira_issues i
JOIN jira_projects p ON i.project_id = p.id
GROUP BY p.project_category_name
ORDER BY nb_issues DESC;

-- ================================================================
-- REQUETE DE CONTROLE : Projets sans issues
-- ================================================================

SELECT
    p.key AS project_key,
    p.name AS project_name,
    p.project_category_name AS categorie
FROM jira_projects p
LEFT JOIN jira_issues i ON i.project_id = p.id
WHERE i.key IS NULL
ORDER BY p.project_category_name, p.key;

"@
[void]$sqlContent.Append($controlSql)

# Ecrire le fichier SQL
Write-FileUtf8 -path $sqlFile -content $sqlContent.ToString()
Log "  Fichier SQL -> $sqlFile"

# ============================================================
# ETAPE 7 : MODE CORRECTION (si demande)
# ============================================================

if ($Mode -eq "Correction") {
  Log "=== ETAPE 7 : Mode correction ==="

  # 7a. Projets orphelins a reassigner
  $orphelins = $allProjects | Where-Object { -not $_.category }
  if ($orphelins.Count -eq 0) {
    Log "  Aucun projet orphelin a corriger."
  } else {
    Log ("  {0} projets orphelins a traiter" -f $orphelins.Count)

    # Afficher les categories disponibles
    $catChoices = @{}
    $idx = 1
    foreach ($cat in ($allCategories | Sort-Object name)) {
      $catChoices[[string]$idx] = $cat
      $idx++
    }

    foreach ($proj in $orphelins) {
      Write-Host "`n========================================" -ForegroundColor Cyan
      Write-Host ("Projet orphelin : {0} - {1}" -f $proj.key, $proj.name) -ForegroundColor Yellow
      Write-Host ("Lead : {0} | Type : {1}" -f $proj.lead, $proj.type) -ForegroundColor Gray
      Write-Host "Categorie a assigner :" -ForegroundColor White

      foreach ($k in ($catChoices.Keys | Sort-Object { [int]$_ })) {
        Write-Host ("  {0,2}. {1}" -f $k, $catChoices[$k].name)
      }
      Write-Host "   0. Ignorer ce projet" -ForegroundColor DarkGray

      $choice = Read-Host "Choix (0-$($catChoices.Count))"

      if ($choice -eq "0" -or -not $choice) {
        Log ("  {0} : ignore" -f $proj.key)
        continue
      }

      if ($catChoices.ContainsKey($choice)) {
        $selectedCat = $catChoices[$choice]
        Log ("  {0} : assignation a '{1}' (id={2})" -f $proj.key, $selectedCat.name, $selectedCat.id)

        # Appel API pour mettre a jour la categorie du projet
        $updateBody = @{ projectCategory = @{ id = [string]$selectedCat.id } } | ConvertTo-Json -Depth 3
        $url = "${baseUrl}/rest/api/3/project/{0}" -f $proj.key
        $resp = Invoke-ApiCall -Method "PUT" -Url $url -Headers $jiraHeaders -Body $updateBody

        if ($resp.ok) {
          Log ("  OK : {0} => {1}" -f $proj.key, $selectedCat.name)
        } else {
          Log ("  ERREUR assignation {0} : status={1} {2}" -f $proj.key, $resp.status, $resp.error) "ERROR"
        }

        Start-Sleep -Milliseconds 300
      } else {
        Log ("  {0} : choix invalide, ignore" -f $proj.key) "WARN"
      }
    }
  }

  # 7b. Projets en categorie archive mais pas archives
  $toArchive = $allProjects | Where-Object { $_.dsType -eq "Archive" -and -not $_.archived }
  if ($toArchive.Count -gt 0) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host ("{0} projets en categorie archive mais encore actifs :" -f $toArchive.Count) -ForegroundColor Yellow
    foreach ($proj in $toArchive) {
      Write-Host ("  {0} - {1} (categorie: {2})" -f $proj.key, $proj.name, $proj.category)
    }
    $confirm = Read-Host "`nArchiver ces projets ? (O/N)"
    if ($confirm -match "^[Oo]") {
      foreach ($proj in $toArchive) {
        $url = "${baseUrl}/rest/api/3/project/{0}/archive" -f $proj.key
        $resp = Invoke-ApiCall -Method "POST" -Url $url -Headers $jiraHeaders
        if ($resp.ok) {
          Log ("  Archive : {0}" -f $proj.key)
        } else {
          Log ("  ERREUR archivage {0} : status={1}" -f $proj.key, $resp.status) "WARN"
        }
        Start-Sleep -Milliseconds 300
      }
    }
  }
}

# ============================================================
# RESUME FINAL
# ============================================================

$dsMetierCount = ($datasources.Keys | Where-Object { $_ -notin @("__ORPHELIN__","__ARCHIVE__","TRANSVERSE") }).Count
$totalDatasources = 1 + $dsMetierCount
if ($IncludeArchived -and ($dsArchive | Measure-Object).Count -gt 0) { $totalDatasources++ }
$totalDatasources++

$cTransverse = ($dsTransverse | Measure-Object).Count
$cArchive = ($dsArchive | Measure-Object).Count
$cActive = ($activeProjects | Measure-Object).Count
$cOrphelin = ($dsOrphelin | Measure-Object).Count

Log "============================================"
Log "RESUME"
Log "============================================"
Log "Site : $baseUrl"
Log "Mode : $Mode"
Log ("Projets total : {0}" -f $allProjects.Count)
Log "---"
Log ("Datasource TRANSVERSE : {0} projets (categories: {1})" -f $cTransverse, ($transverseCategories -join ", "))
Log ("Datasources METIER : {0} datasources" -f $dsMetierCount)
foreach ($ds in ($dsMetier | Sort-Object { $_.name })) {
  Log ("  - {0} : {1} projets" -f $ds.name, $ds.count)
}
Log ("Datasource ARCHIVES : {0} projets" -f $cArchive)
Log ("Datasource GLOBALE : {0} projets actifs" -f $cActive)
Log ("Projets orphelins : {0}" -f $cOrphelin)
Log "---"
Log ("Total datasources a creer : {0}" -f $totalDatasources)
Log ("Anomalies detectees : {0}" -f $anomalies.Count)
Log "---"
Log "Fichiers generes :"
Log "  Inventaire CSV : $csvFile"
Log "  Requetes SQL   : $sqlFile"
Log "  Journal        : $logFile"
if ($anomalies.Count -gt 0) { Log "  Coherence      : $coherenceFile" }
Log "============================================"
Log "Termine."
# -------------------- Popup resume --------------------
$summaryLines = @(
  "Gestion Datasources Analytics terminee.",
  "",
  "Site: $baseUrl",
  "Mode: $Mode",
  ("Projets : {0}" -f $allProjects.Count),
  "",
  ("Datasources a creer : {0}" -f $totalDatasources),
  ("  TRANSVERSE : {0} projets" -f $cTransverse),
  ("  METIER : {0} datasources" -f $dsMetierCount),
  ("  ARCHIVES : {0} projets" -f $cArchive),
  ("  GLOBALE : {0} projets" -f $cActive),
  "",
  ("Anomalies : {0}" -f $anomalies.Count),
  "",
  "Fichier SQL : $sqlFile"
)
$summaryMsg = $summaryLines -join "`n"

# Creer une fenetre invisible TopMost pour forcer la popup au premier plan
$topForm = New-Object System.Windows.Forms.Form
$topForm.TopMost = $true
$topForm.ShowInTaskbar = $false
$topForm.Size = New-Object System.Drawing.Size(1,1)
$topForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$topForm.Location = New-Object System.Drawing.Point(-100,-100)
$topForm.Show()
$topForm.Hide()

[System.Windows.Forms.MessageBox]::Show(
  $topForm,
  $summaryMsg,
  "Datasources Analytics",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

$topForm.Dispose()