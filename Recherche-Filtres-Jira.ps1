<#
.SYNOPSIS
  Recherche-Filtres-Jira.ps1
  Recherche dans tous les filtres Jira ceux dont le JQL contient
  des references a des champs Tempo qui changent avec la migration Forge,
  et propose le JQL corrige.

.DESCRIPTION
  Ce script :
  1. Parcourt tous les filtres Jira de l'instance
  2. Detecte les references aux champs qui cassent avec la migration Forge :
     - issue.property[tempo-team] => remplacer par "Tempo Team" (cf[10031])
     - issue.property[tempo-account] => remplacer par Account (cf[10032])
     - issue.internal.* => a verifier manuellement
  3. NE TOUCHE PAS aux champs Account et Tempo Team utilises normalement
     (ils restent valides apres migration Forge)
  4. Propose un JQL corrige et peut l'appliquer avec -Execute

  CHAMPS JIRADOT :
    Account    = customfield_10032 (io.tempo.jira__account) => VALIDE apres Forge
    Tempo Team = customfield_10031 (io.tempo.jira__team)    => VALIDE apres Forge
    Team       = customfield_10001 (atlassian-team)         => Natif Jira

  CE QUI CASSE AVEC FORGE :
    issue.property[tempo-team].*    => Remplacer par "Tempo Team"
    issue.property[tempo-account].* => Remplacer par Account
    issue.property[tempo-*].*       => A verifier
    issue.internal.*                => Obsolete

  API :
    GET  /rest/api/3/filter/search?expand=jql,viewUrl,searchUrl,owner
    PUT  /rest/api/3/filter/{id}

  CREDENTIALS :
    secrets\site-admin.xml => Jiradot

  FICHIERS GENERES (dans exports\) :
    MigrationFiltres_{ts}.csv  => Rapport detaille
    MigrationFiltres_{ts}.log  => Journal d execution

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Recherche-Filtres-Jira.ps1              # Dry-run
  .\Recherche-Filtres-Jira.ps1 -Execute     # Applique les corrections
#>

[CmdletBinding()]
param(
  [switch] $Execute,
  [int]    $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$modeLabel = if ($Execute) { "EXECUTION" } else { "DRY-RUN" }

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("MigrationFiltres_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("MigrationFiltres_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# REGLES DE CORRECTION JQL
#
# IMPORTANT : les champs "Account" et "Tempo Team" utilises
# normalement dans les JQL (ex: Account = "Projet Alpha")
# restent VALIDES apres migration Forge.
#
# Seules les PROPRIETES D'ISSUE (issue.property[...]) et
# les CHAMPS INTERNES (issue.internal) cassent.
# ============================================================

$correctionRules = @(

  # --- ISSUE.PROPERTY TEMPO TEAM ---
  # issue.property[tempo-team].name = "..." => "Tempo Team" = "..."
  @{
    Name        = "issue.property[tempo-team].name = valeur"
    Pattern     = 'issue\.property\[tempo-team\]\.name\s*=\s*"([^"]*)"'
    Replacement = '"Tempo Team" = "$1"'
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team => champ custom Tempo Team (cf[10031])"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-team].name != valeur"
    Pattern     = 'issue\.property\[tempo-team\]\.name\s*!=\s*"([^"]*)"'
    Replacement = '"Tempo Team" != "$1"'
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team != => champ custom Tempo Team"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-team].name in (...)"
    Pattern     = 'issue\.property\[tempo-team\]\.name\s+in\s*(\([^)]*\))'
    Replacement = '"Tempo Team" in $1'
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team in => champ custom Tempo Team"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-team].name not in (...)"
    Pattern     = 'issue\.property\[tempo-team\]\.name\s+not\s+in\s*(\([^)]*\))'
    Replacement = '"Tempo Team" not in $1'
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team not in => champ custom Tempo Team"
    Severity    = "Haute"
    AutoFix     = $true
  },
  # Variantes avec .id ou autre attribut
  @{
    Name        = "issue.property[tempo-team].id = valeur"
    Pattern     = 'issue\.property\[tempo-team\]\.id\s*=\s*(\S+)'
    Replacement = '"Tempo Team" = $1'
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team par ID => champ custom Tempo Team"
    Severity    = "Haute"
    AutoFix     = $true
  },
  # Catch-all tempo-team property
  @{
    Name        = "issue.property[tempo-team] generique"
    Pattern     = 'issue\.property\[tempo-team\][^\s]*\s*(?:=|!=|~|!~|in|not\s+in|is)\s*(?:"[^"]*"|\([^)]*\)|EMPTY|not\s+EMPTY|\S+)'
    Replacement = ""
    Category    = "Tempo Team Property"
    Description = "Propriete issue Tempo Team non reconnue - suppression (verifier manuellement)"
    Severity    = "Haute"
    AutoFix     = $false
  },

  # --- ISSUE.PROPERTY TEMPO ACCOUNT ---
  # issue.property[tempo-account].name = "..." => Account = "..."
  @{
    Name        = "issue.property[tempo-account].name = valeur"
    Pattern     = 'issue\.property\[tempo-account\]\.name\s*=\s*"([^"]*)"'
    Replacement = 'Account = "$1"'
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account => champ custom Account (cf[10032])"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-account].name != valeur"
    Pattern     = 'issue\.property\[tempo-account\]\.name\s*!=\s*"([^"]*)"'
    Replacement = 'Account != "$1"'
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account != => champ custom Account"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-account].name in (...)"
    Pattern     = 'issue\.property\[tempo-account\]\.name\s+in\s*(\([^)]*\))'
    Replacement = 'Account in $1'
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account in => champ custom Account"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-account].name not in (...)"
    Pattern     = 'issue\.property\[tempo-account\]\.name\s+not\s+in\s*(\([^)]*\))'
    Replacement = 'Account not in $1'
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account not in => champ custom Account"
    Severity    = "Haute"
    AutoFix     = $true
  },
  @{
    Name        = "issue.property[tempo-account].key = valeur"
    Pattern     = 'issue\.property\[tempo-account\]\.key\s*=\s*"([^"]*)"'
    Replacement = 'Account = "$1"'
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account par key => champ custom Account"
    Severity    = "Haute"
    AutoFix     = $true
  },
  # Catch-all tempo-account property
  @{
    Name        = "issue.property[tempo-account] generique"
    Pattern     = 'issue\.property\[tempo-account\][^\s]*\s*(?:=|!=|~|!~|in|not\s+in|is)\s*(?:"[^"]*"|\([^)]*\)|EMPTY|not\s+EMPTY|\S+)'
    Replacement = ""
    Category    = "Tempo Account Property"
    Description = "Propriete issue Tempo Account non reconnue - suppression (verifier manuellement)"
    Severity    = "Haute"
    AutoFix     = $false
  },

  # --- AUTRES ISSUE.PROPERTY TEMPO ---
  @{
    Name        = "issue.property[tempo-*] autre"
    Pattern     = 'issue\.property\[tempo-[^\]]+\][^\s]*\s*(?:=|!=|~|!~|in|not\s+in|is)\s*(?:"[^"]*"|\([^)]*\)|EMPTY|not\s+EMPTY|\S+)'
    Replacement = ""
    Category    = "Tempo Property Autre"
    Description = "Propriete issue Tempo inconnue - a verifier manuellement"
    Severity    = "Moyenne"
    AutoFix     = $false
  },

  # --- ISSUE.INTERNAL ---
  @{
    Name        = "issue.internal generique"
    Pattern     = 'issue\.internal\.[^\s]+\s*(?:=|!=|~|!~|in|not\s+in|is)\s*(?:"[^"]*"|\([^)]*\)|EMPTY|not\s+EMPTY|\S+)'
    Replacement = ""
    Category    = "Issue Internal"
    Description = "Champ interne issue - obsolete apres migration"
    Severity    = "Haute"
    AutoFix     = $false
  },

  # --- ISSUE.PROPERTY GENERIQUE (non-Tempo) ---
  @{
    Name        = "issue.property generique"
    Pattern     = 'issue\.property\[[^\]]+\][^\s]*\s*(?:=|!=|~|!~|in|not\s+in|is)\s*(?:"[^"]*"|\([^)]*\)|EMPTY|not\s+EMPTY|\S+)'
    Replacement = ""
    Category    = "Issue Property Autre"
    Description = "Propriete issue non-Tempo - a verifier manuellement"
    Severity    = "Moyenne"
    AutoFix     = $false
  }
)

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
# FONCTION DE CORRECTION JQL
# ============================================================

function Correct-JQL([string]$jql) {
  $correctedJql = $jql
  $appliedRules = @()
  $hasAutoFix   = $true

  foreach ($rule in $correctionRules) {
    if ($correctedJql -match $rule.Pattern) {
      $matches = [regex]::Matches($correctedJql, $rule.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      foreach ($m in $matches) {
        $appliedRules += @{
          RuleName    = $rule.Name
          Category    = $rule.Category
          Severity    = $rule.Severity
          MatchedText = $m.Value
          Description = $rule.Description
          AutoFix     = $rule.AutoFix
        }
      }

      if ($rule.Replacement -eq "") {
        # Supprimer la clause
        $correctedJql = [regex]::Replace($correctedJql, $rule.Pattern, "##REMOVED##", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $hasAutoFix = $false
      } else {
        # Remplacer par l'equivalent
        $correctedJql = [regex]::Replace($correctedJql, $rule.Pattern, $rule.Replacement, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      }
    }
  }

  # Nettoyer les ##REMOVED## et les AND/OR orphelins
  if ($correctedJql.Contains("##REMOVED##")) {
    $correctedJql = [regex]::Replace($correctedJql, '\s+AND\s+##REMOVED##', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = [regex]::Replace($correctedJql, '##REMOVED##\s+AND\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = [regex]::Replace($correctedJql, '\s+OR\s+##REMOVED##', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = [regex]::Replace($correctedJql, '##REMOVED##\s+OR\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = $correctedJql -replace '##REMOVED##', ''
    $correctedJql = [regex]::Replace($correctedJql, '\(\s*\)', '')
    $correctedJql = [regex]::Replace($correctedJql, '\s+', ' ')
    $correctedJql = [regex]::Replace($correctedJql, '^\s*(AND|OR)\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = [regex]::Replace($correctedJql, '\s+(AND|OR)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = [regex]::Replace($correctedJql, '\s+(AND|OR)\s+(ORDER\s+BY)', ' $2', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $correctedJql = $correctedJql.Trim()
  }

  # Determiner le type d'action
  $allAutoFix = ($appliedRules | Where-Object { -not $_.AutoFix } | Measure-Object).Count -eq 0
  $isEmptyAfter = ([string]::IsNullOrWhiteSpace($correctedJql) -or $correctedJql -match '^\s*(ORDER\s+BY\s+.*)?$')

  return @{
    Original     = $jql
    Corrected    = $correctedJql
    HasChanges   = ($correctedJql -ne $jql)
    AppliedRules = $appliedRules
    AllAutoFix   = $allAutoFix
    IsEmpty      = $isEmptyAfter
  }
}

# ============================================================
# DEBUT
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  MIGRATION FILTRES JIRA (TEMPO FORGE)" -ForegroundColor Cyan
Write-Host ("  Mode : {0}" -f $modeLabel) -ForegroundColor $(if ($Execute) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if (-not $Execute) {
  Write-Host "  Mode DRY-RUN : aucun filtre ne sera modifie." -ForegroundColor Green
  Write-Host "  Relancez avec -Execute pour appliquer les corrections." -ForegroundColor Green
  Write-Host ""
}

Write-Host "  Champs qui CASSENT avec la migration Forge :" -ForegroundColor White
Write-Host "    - issue.property[tempo-team]    => remplace par 'Tempo Team' (cf[10031])" -ForegroundColor DarkGray
Write-Host "    - issue.property[tempo-account] => remplace par 'Account' (cf[10032])" -ForegroundColor DarkGray
Write-Host "    - issue.internal.*              => obsolete, supprime" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Champs qui RESTENT VALIDES :" -ForegroundColor White
Write-Host "    - Account = '...'               => OK (cf[10032] inchange)" -ForegroundColor Green
Write-Host "    - 'Tempo Team' = '...'          => OK (cf[10031] inchange)" -ForegroundColor Green
Write-Host ""

Log "================================================================"
Log ("  MIGRATION FILTRES JIRA - {0}" -f $modeLabel)
Log "================================================================"

$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
Log ("  Site : {0}" -f $site.BaseUrl)

$startTime = Get-Date

# ============================================================
# PARCOURS DES FILTRES
# ============================================================

Log "=== Parcours des filtres ==="

$startAt    = 0
$maxResults = 100
$finished   = $false
$cFilters   = 0
$cMatches   = 0
$cAutoFix   = 0
$cManual    = 0
$cEmptyAfter= 0
$cPages     = 0
$results    = New-Object System.Collections.Generic.List[object]

while (-not $finished) {
  $cPages++

  $url = "{0}/rest/api/3/filter/search?expand=jql,viewUrl,searchUrl,owner&startAt={1}&maxResults={2}" -f $site.BaseUrl, $startAt, $maxResults
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

  if (-not $resp.ok) {
    Log ("  Erreur page {0} : status={1}" -f $cPages, $resp.status) "ERROR"
    break
  }

  $json = $resp.content | ConvertFrom-Json
  $filters = $json.values
  $total   = [int]$json.total
  $fc      = ($filters | Measure-Object).Count

  if ($fc -eq 0) { break }

  foreach ($filter in $filters) {
    $cFilters++
    $filterName = [string]$filter.name
    $filterId   = [string]$filter.id
    $filterJql  = if ($filter.jql) { [string]$filter.jql } else { "" }
    $filterOwner = ""
    if ($filter.owner -and $filter.owner.displayName) {
      $filterOwner = [string]$filter.owner.displayName
    }
    $filterOwnerAccountId = ""
    if ($filter.owner -and $filter.owner.accountId) {
      $filterOwnerAccountId = [string]$filter.owner.accountId
    }
    $filterViewUrl = if ($filter.viewUrl) { [string]$filter.viewUrl } else { "" }

    # Appliquer les corrections
    $correction = Correct-JQL $filterJql

    if ($correction.HasChanges) {
      $cMatches++

      $ruleNames  = ($correction.AppliedRules | ForEach-Object { $_.RuleName }) -join ", "
      $categories = ($correction.AppliedRules | ForEach-Object { $_.Category } | Sort-Object -Unique) -join ", "
      $severities = ($correction.AppliedRules | ForEach-Object { $_.Severity } | Sort-Object -Unique) -join ", "
      $matchTexts = ($correction.AppliedRules | ForEach-Object { $_.MatchedText }) -join " | "

      if ($correction.IsEmpty) {
        $action = "VIDE APRES CORRECTION"
        $cEmptyAfter++
      } elseif ($correction.AllAutoFix) {
        $action = "AUTO-CORRIGEABLE"
        $cAutoFix++
      } else {
        $action = "VERIFICATION MANUELLE"
        $cManual++
      }

      Write-Host "-----------------------------------" -ForegroundColor DarkGray
      Write-Host ("  Filtre   : {0} (ID:{1})" -f $filterName, $filterId) -ForegroundColor Yellow
      Write-Host ("  Owner    : {0}" -f $filterOwner) -ForegroundColor DarkGray
      Write-Host ("  Categorie: {0}" -f $categories) -ForegroundColor DarkGray
      Write-Host ("  Action   : {0}" -f $action) -ForegroundColor $(
        if ($action -eq "AUTO-CORRIGEABLE") { "Green" }
        elseif ($action -like "*VIDE*") { "Red" }
        else { "Yellow" }
      )
      Write-Host ("  Clauses  : {0}" -f $matchTexts) -ForegroundColor Red
      Write-Host ("  JQL avant: {0}" -f $correction.Original) -ForegroundColor DarkGray
      if ($correction.IsEmpty) {
        Write-Host ("  JQL apres: (VIDE - filtre a supprimer ou reconstruire)") -ForegroundColor Red
      } else {
        Write-Host ("  JQL apres: {0}" -f $correction.Corrected) -ForegroundColor Green
      }

      Log ("  MATCH #{0} : '{1}' (ID:{2}) owner={3} action={4} regles=[{5}]" -f $cMatches, $filterName, $filterId, $filterOwner, $action, $ruleNames)

      $results.Add(@{
        FilterName     = $filterName
        FilterId       = $filterId
        Owner          = $filterOwner
        OwnerAccountId = $filterOwnerAccountId
        ViewUrl        = $filterViewUrl
        JqlOriginal    = $correction.Original
        JqlCorrected   = $correction.Corrected
        IsEmpty        = $correction.IsEmpty
        AllAutoFix     = $correction.AllAutoFix
        Categories     = $categories
        Severities     = $severities
        RuleNames      = $ruleNames
        MatchedClauses = $matchTexts
        Action         = $action
        Status         = "PLANIFIE"
      }) | Out-Null
    }
  }

  $startAt += $json.maxResults
  $finished = $startAt -ge $total

  Write-Progress -Activity "Analyse des filtres" `
    -Status ("{0}/{1} filtres, {2} a traiter" -f $cFilters, $total, $cMatches) `
    -PercentComplete ([int](100 * [Math]::Min($startAt, $total) / [Math]::Max($total, 1)))

  Start-Sleep -Milliseconds 150
}

Write-Progress -Activity "Analyse" -Completed

Log ("  Filtres parcourus     : {0}" -f $cFilters)
Log ("  Filtres a traiter     : {0}" -f $cMatches)
Log ("  Auto-corrigeables     : {0}" -f $cAutoFix)
Log ("  Verification manuelle : {0}" -f $cManual)
Log ("  Vides apres correction: {0}" -f $cEmptyAfter)

# ============================================================
# EXPORT CSV
# ============================================================

Log "=== Export CSV ==="

$csvCols = @("FilterId","FilterName","Owner","Action","Categories","Severities","MatchedClauses","JqlOriginal","JqlCorrected","Status","ViewUrl")
$csvHeaderLine = ($csvCols | ForEach-Object { CsvEscape $_ }) -join ";"
Write-CsvToFile $csvFile $csvHeaderLine -Header

foreach ($r in ($results | Sort-Object { $_.Action }, { $_.FilterName })) {
  $vals = @(
    $r.FilterId, $r.FilterName, $r.Owner, $r.Action, $r.Categories, $r.Severities,
    $r.MatchedClauses, $r.JqlOriginal, $r.JqlCorrected, $r.Status, $r.ViewUrl
  )
  $line = ($vals | ForEach-Object { CsvEscape $_ }) -join ";"
  Write-CsvToFile $csvFile $line
}

Log ("  CSV -> {0}" -f $csvFile)

# ============================================================
# EXECUTION (si -Execute)
# ============================================================

$cOK = 0; $cKO = 0; $cSkip = 0

if ($Execute -and $cMatches -gt 0) {
  Write-Host ""
  Write-Host "========================================" -ForegroundColor Red
  Write-Host "  EXECUTION : CORRECTION DES FILTRES" -ForegroundColor Red
  Write-Host "========================================" -ForegroundColor Red
  Write-Host ""
  Write-Host ("  {0} filtres a traiter :" -f $cMatches) -ForegroundColor Red
  Write-Host ("    - {0} auto-corrigeables (seront modifies)" -f $cAutoFix) -ForegroundColor Green
  Write-Host ("    - {0} verification manuelle (seront IGNORES)" -f $cManual) -ForegroundColor Yellow
  Write-Host ("    - {0} vides apres correction (seront IGNORES)" -f $cEmptyAfter) -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Seuls les filtres AUTO-CORRIGEABLES seront modifies." -ForegroundColor White
  Write-Host "  Les autres necessitent une intervention manuelle." -ForegroundColor DarkGray
  Write-Host ""
  $confirm = Read-Host "  Tapez CORRIGER pour confirmer (ou autre chose pour annuler)"

  if ($confirm -eq "CORRIGER") {
    Log "=== Execution des corrections ==="

    foreach ($r in $results) {
      # Ne corriger que les auto-fix
      if (-not $r.AllAutoFix -or $r.IsEmpty) {
        $skipReason = if ($r.IsEmpty) { "vide apres correction" } else { "verification manuelle requise" }
        $r.Status = "IGNORE ({0})" -f $skipReason
        $cSkip++
        Log ("  IGNORE '{0}' (ID:{1}) - {2}" -f $r.FilterName, $r.FilterId, $skipReason)
        Write-Host ("    SKIP {0,-35} (ID:{1}) - {2}" -f $r.FilterName, $r.FilterId, $skipReason) -ForegroundColor Yellow
        continue
      }

      Log ("  Correction '{0}' (ID:{1})..." -f $r.FilterName, $r.FilterId)

      $updateBody = @{
        jql = $r.JqlCorrected
      } | ConvertTo-Json -Depth 5 -Compress

      $updateUrl = "{0}/rest/api/3/filter/{1}" -f $site.BaseUrl, $r.FilterId
      $updateResp = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $site.Headers -Body $updateBody

      if ($updateResp.ok) {
        $r.Status = "OK"
        $cOK++
        Log ("    OK : filtre corrige")
        Write-Host ("    OK  {0,-35} (ID:{1})" -f $r.FilterName, $r.FilterId) -ForegroundColor Green
      } else {
        $errStatus = "ERREUR:{0}" -f $updateResp.status
        $r.Status = $errStatus
        $cKO++
        Log ("    ERREUR : status={0}" -f $updateResp.status) "ERROR"
        Write-Host ("    ERR {0,-35} (ID:{1}) status={2}" -f $r.FilterName, $r.FilterId, $updateResp.status) -ForegroundColor Red

        if ($updateResp.body) {
          try {
            $errJson = $updateResp.body | ConvertFrom-Json
            if ($errJson.errorMessages) {
              foreach ($em in $errJson.errorMessages) {
                Log ("      Detail : {0}" -f $em) "ERROR"
                Write-Host ("      -> {0}" -f $em) -ForegroundColor DarkGray
              }
            }
            if ($errJson.errors) {
              $errJson.errors.PSObject.Properties | ForEach-Object {
                Log ("      Champ {0} : {1}" -f $_.Name, $_.Value) "ERROR"
                Write-Host ("      -> {0}: {1}" -f $_.Name, $_.Value) -ForegroundColor DarkGray
              }
            }
          } catch {}
        }
      }

      Start-Sleep -Milliseconds 300
    }

    # Mise a jour du CSV
    Write-CsvToFile $csvFile $csvHeaderLine -Header
    foreach ($r in ($results | Sort-Object { $_.Action }, { $_.FilterName })) {
      $vals = @(
        $r.FilterId, $r.FilterName, $r.Owner, $r.Action, $r.Categories, $r.Severities,
        $r.MatchedClauses, $r.JqlOriginal, $r.JqlCorrected, $r.Status, $r.ViewUrl
      )
      $line = ($vals | ForEach-Object { CsvEscape $_ }) -join ";"
      Write-CsvToFile $csvFile $line
    }

  } else {
    Log "  Execution annulee par l'utilisateur."
    Write-Host "  Annule." -ForegroundColor Yellow
    foreach ($r in $results) { $r.Status = "ANNULE" }
  }
}

# ============================================================
# RESULTATS
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESULTATS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($cMatches -eq 0) {
  Write-Host "  Aucun filtre ne contient de references Tempo obsoletes." -ForegroundColor Green
  Write-Host "  Les champs Account et Tempo Team sont utilises normalement" -ForegroundColor Green
  Write-Host "  et restent valides apres migration Forge." -ForegroundColor Green
} else {
  # Stats par action
  Write-Host "  Par type d'action :" -ForegroundColor White
  $actionGroups = $results | Group-Object { $_.Action } | Sort-Object Name
  foreach ($ag in $actionGroups) {
    $agCount = ($ag.Group | Measure-Object).Count
    $color = switch -Wildcard ($ag.Name) {
      "AUTO*"    { "Green" }
      "VIDE*"    { "Red" }
      "VERIF*"   { "Yellow" }
      default    { "White" }
    }
    Write-Host ("    {0,-35} : {1} filtres" -f $ag.Name, $agCount) -ForegroundColor $color
  }

  Write-Host ""

  # Stats par categorie
  Write-Host "  Par categorie :" -ForegroundColor White
  $catGroups = $results | ForEach-Object {
    $_.Categories -split ", " | ForEach-Object { $_.Trim() }
  } | Group-Object | Sort-Object Count -Descending
  foreach ($cg in $catGroups) {
    Write-Host ("    {0,-35} : {1}" -f $cg.Name, $cg.Count) -ForegroundColor Yellow
  }

  Write-Host ""

  # Stats par owner
  Write-Host "  Par owner (top 15) :" -ForegroundColor White
  $ownerGroups = $results | Group-Object { $_.Owner } | Sort-Object Count -Descending | Select-Object -First 15
  foreach ($og in $ownerGroups) {
    $ogCount = ($og.Group | Measure-Object).Count
    Write-Host ("    {0,-35} : {1} filtres" -f $og.Name, $ogCount) -ForegroundColor Yellow
  }
}

# ============================================================
# RESUME
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  RESUME ({0})" -f $modeLabel) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Filtres parcourus     : {0}" -f $cFilters)
Write-Host ("  Filtres a traiter     : {0}" -f $cMatches) -ForegroundColor $(if ($cMatches -gt 0) { "Yellow" } else { "Green" })
Write-Host ("    - Auto-corrigeables : {0}" -f $cAutoFix) -ForegroundColor $(if ($cAutoFix -gt 0) { "Green" } else { "DarkGray" })
Write-Host ("    - Verif. manuelle   : {0}" -f $cManual) -ForegroundColor $(if ($cManual -gt 0) { "Yellow" } else { "DarkGray" })
Write-Host ("    - Vides apres       : {0}" -f $cEmptyAfter) -ForegroundColor $(if ($cEmptyAfter -gt 0) { "Red" } else { "DarkGray" })

if ($Execute) {
  Write-Host ""
  Write-Host ("  Corriges              : {0}" -f $cOK) -ForegroundColor Green
  Write-Host ("  Echoues               : {0}" -f $cKO) -ForegroundColor $(if ($cKO -gt 0) { "Red" } else { "Green" })
  Write-Host ("  Ignores               : {0}" -f $cSkip) -ForegroundColor $(if ($cSkip -gt 0) { "Yellow" } else { "DarkGray" })
}

Write-Host ""
Write-Host ("  Duree                 : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  CSV                   : {0}" -f $csvFile)
Write-Host ("  LOG                   : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

if (-not $Execute -and $cAutoFix -gt 0) {
  Write-Host ""
  Write-Host "  Pour appliquer les auto-corrections :" -ForegroundColor Yellow
  Write-Host "    .\Recherche-Filtres-Jira.ps1 -Execute" -ForegroundColor Yellow
}

Log "Termine."