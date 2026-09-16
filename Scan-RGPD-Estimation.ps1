<#
.SYNOPSIS
  Scan-RGPD-Estimation.ps1 v1.13.2
  Estimation statistique du nombre de tickets contenant des donnees sensibles
  dans un ou plusieurs projets Jira (RGPD).

.DESCRIPTION
  Methode : Echantillonnage stratifie par statut Done/NotDone + formule de Cochran
  IC      : 95 % par defaut, marge +/- 2.5 %
  Par projet  : taille d echantillon calculee individuellement
  Global      : somme ponderee des estimations par projet
  Petit projet (<n0 tickets) : scan exhaustif
  Grand projet : echantillon stratifie

  NOUVEAUTES v1.13.2 :
    - Get-ProjectCountAndLastUpdated : fusionne comptage total + date MAJ en UNE seule passe
    - Comptage Done en une seule passe (NotDone = Total - Done, jamais de 3e pagination)
    - Tentative GET api/2 maxResults=1 (retourne total si endpoint actif)
    - Suppression limite 10 pages : pagination complete (progression toutes les 20 pages)
    - Log du temps ecoule par projet

  CREDENTIALS :
    secrets\site-admin.xml => @{ SiteUrl; Email; ApiTokenSecureString }

.NOTES
  Auteur : Frederic GUEDJ
  Version: 1.13.2

.EXAMPLE
  .\Scan-RGPD-Estimation.ps1 -Projects "*"
  .\Scan-RGPD-Estimation.ps1 -Projects "CPT,CNT,PACT"
  .\Scan-RGPD-Estimation.ps1 -MarginOfError 0.03
#>

[CmdletBinding()]
param(
  [string]$Projects = "",
  [double]$MarginOfError = 0.025,
  [double]$ConfidenceLevel = 0.95,
  [int]$MaxRetries = 5
)

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir = if ($PSScriptRoot) {
  $PSScriptRoot
} elseif ($psISE) {
  Split-Path $psISE.CurrentFile.FullPath
} else {
  Split-Path $MyInvocation.MyCommand.Path
}

$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("RGPD-Estimation_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("RGPD-Estimation_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level = "INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  if ($logFile) { Add-Content -Path $logFile -Value $line -Encoding UTF8 }
}

function Add-CsvLine([string]$line) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::AppendAllText($csvFile, "$line`r`n", $utf8Bom)
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
  param(
    [string]$Method,
    [string]$Url,
    [hashtable]$Headers,
    [string]$Body = $null
  )
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
      }
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
      return @{ ok = $true; status = [int]$resp.StatusCode; content = $contentUtf8 }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $rd = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $rd.ReadToEnd(); $rd.Close()
      } catch {}
      if ($attempt -gt $MaxRetries) {
        return @{ ok = $false; status = $status; error = $_.Exception.Message; body = $errBody }
      }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(5, $attempt)))
        Log ("  Retry {0} status={1} in {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }
      return @{ ok = $false; status = $status; error = $_.Exception.Message; body = $errBody }
    }
  }
}

# ============================================================
# CREDENTIALS
# ============================================================

function Import-SiteCredentials {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
  param([string]$XmlPath)

  if (-not (Test-Path $XmlPath)) {
    Log "Fichier credentials introuvable : $XmlPath" "ERROR"
    throw "Credentials manquants"
  }
  $data  = Import-Clixml -Path $XmlPath
  $url   = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
  }
}

# ============================================================
# JIRA API : SEARCH (POST /rest/api/3/search/jql)
# ============================================================

function Invoke-JqlSearch {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$Jql,
    [int]$MaxResults = 50,
    [string[]]$Fields = @("key"),
    [string]$NextPageToken = $null
  )

  $bodyObj = @{
    jql        = $Jql
    maxResults = $MaxResults
    fields     = $Fields
  }
  if ($NextPageToken) {
    $bodyObj["nextPageToken"] = $NextPageToken
  }

  $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress
  $url = "$BaseUrl/rest/api/3/search/jql"

  $resp = Invoke-ApiCall -Method "POST" -Url $url -Headers $Headers -Body $bodyJson

  # Fallback : si POST echoue, essayer GET api/2
  if (-not $resp.ok) {
    if ($resp.status -ne 410) {
      Log ("  POST /search/jql echoue (status={0}), fallback GET api/2" -f $resp.status) "WARN"
    }
    $fieldsStr = $Fields -join ","
    $encodedJql = [System.Uri]::EscapeDataString($Jql)
    $getUrl = "{0}/rest/api/2/search?jql={1}&maxResults={2}&fields={3}" -f $BaseUrl, $encodedJql, $MaxResults, $fieldsStr
    $resp = Invoke-ApiCall -Method "GET" -Url $getUrl -Headers $Headers
  }

  return $resp
}

# ============================================================
# COMPTAGE + DERNIERE MAJ EN UNE SEULE PASSE
# Strategie :
#   1. GET api/2 maxResults=1 ORDER BY updated DESC -> total + lastUpdated (1 appel)
#   2. Si 410 : POST api/3 pour lastUpdated (1er ticket) + pagination curseur pour compter
# ============================================================

function Get-ProjectCountAndLastUpdated {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$ProjectKey,
    [string]$ExtraJql = ""
  )

  $jql = "project = `"$ProjectKey`""
  if ($ExtraJql) { $jql += " AND $ExtraJql" }
  $jqlOrdered = "$jql ORDER BY updated DESC"

  # --- Tentative rapide : GET api/2 avec maxResults=1 ---
  $encodedJql = [System.Uri]::EscapeDataString($jqlOrdered)
  $getUrl = "{0}/rest/api/2/search?jql={1}&maxResults=1&fields=updated" -f $BaseUrl, $encodedJql
  $resp = Invoke-ApiCall -Method "GET" -Url $getUrl -Headers $Headers

  if ($resp.ok) {
    $json = $resp.content | ConvertFrom-Json
    if ($null -ne $json.total) {
      # Succes : on a total + lastUpdated en 1 seul appel
      $lastUpd = "(aucun ticket)"
      if ($json.issues -and ($json.issues | Measure-Object).Count -gt 0) {
        $updRaw = $json.issues[0].fields.updated
        if ($updRaw) {
          try { $lastUpd = ([datetime]::Parse($updRaw)).ToString("yyyy-MM-dd HH:mm") }
          catch { $lastUpd = [string]$updRaw.Substring(0, [Math]::Min(16, $updRaw.Length)) }
        }
      }
      return @{ ok = $true; total = [long]$json.total; lastUpdated = $lastUpd }
    }
  }

  # --- Fallback : POST api/3 (pagination curseur complete) ---
  # Premier appel : ORDER BY updated DESC pour avoir la date MAJ
  $respFirst = Invoke-JqlSearch -BaseUrl $BaseUrl -Headers $Headers -Jql $jqlOrdered -MaxResults 100 -Fields @("key", "updated")

  if (-not $respFirst.ok) {
    return @{ ok = $false; total = 0; lastUpdated = "(erreur)"; status = $respFirst.status }
  }

  $jsonFirst = $respFirst.content | ConvertFrom-Json

  # Recuperer lastUpdated depuis le premier ticket
  $lastUpd = "(aucun ticket)"
  if ($jsonFirst.issues -and ($jsonFirst.issues | Measure-Object).Count -gt 0) {
    $updRaw = $jsonFirst.issues[0].fields.updated
    if ($updRaw) {
      try { $lastUpd = ([datetime]::Parse($updRaw)).ToString("yyyy-MM-dd HH:mm") }
      catch { $lastUpd = [string]$updRaw.Substring(0, [Math]::Min(16, $updRaw.Length)) }
    }
  }

  # Si total est dans la reponse (fallback api/2 a fonctionne)
  if ($null -ne $jsonFirst.total) {
    return @{ ok = $true; total = [long]$jsonFirst.total; lastUpdated = $lastUpd }
  }

  # Sinon : pagination curseur COMPLETE pour compter (sans limite de pages)
  Log ("    Comptage par pagination pour {0}..." -f $ProjectKey)
  $count = 0
  if ($jsonFirst.issues) { $count = ($jsonFirst.issues | Measure-Object).Count }
  $nextToken = $jsonFirst.nextPageToken
  $isLast = $jsonFirst.isLast
  $pageNum = 1

  # Pour le comptage on utilise ORDER BY key ASC (plus rapide, pas besoin de updated)
  $jqlCount = "$jql ORDER BY key ASC"

  # Si la premiere page est deja la derniere
  if ($isLast -or -not $nextToken) {
    return @{ ok = $true; total = [long]$count; lastUpdated = $lastUpd }
  }

  # Continuer la pagination avec key ASC et fields=key (leger)
  # On recommence depuis le debut en key ASC pour un comptage propre
  $count = 0
  $nextToken = $null
  $isLast = $false
  $pageNum = 0

  do {
    $pageNum++
    $resp = Invoke-JqlSearch -BaseUrl $BaseUrl -Headers $Headers -Jql $jqlCount -MaxResults 100 -Fields @("key") -NextPageToken $nextToken
    if (-not $resp.ok) { break }
    $jsonPage = $resp.content | ConvertFrom-Json
    if ($jsonPage.issues) { $count += ($jsonPage.issues | Measure-Object).Count }
    $isLast = $jsonPage.isLast
    $nextToken = $jsonPage.nextPageToken

    if ($pageNum % 20 -eq 0) {
      Log ("    Comptage {0} : {1} tickets (page {2})..." -f $ProjectKey, $count, $pageNum)
    }
    Start-Sleep -Milliseconds 100
  } while (-not $isLast -and $nextToken)

  return @{ ok = $true; total = [long]$count; lastUpdated = $lastUpd }
}

# ============================================================
# COMPTAGE DONE SEUL (une seule passe)
# ============================================================

function Get-DoneCount {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$ProjectKey
  )

  $jql = "project = `"$ProjectKey`" AND statusCategory = Done"

  # Tentative rapide : GET api/2
  $encodedJql = [System.Uri]::EscapeDataString($jql)
  $getUrl = "{0}/rest/api/2/search?jql={1}&maxResults=1&fields=key" -f $BaseUrl, $encodedJql
  $resp = Invoke-ApiCall -Method "GET" -Url $getUrl -Headers $Headers

  if ($resp.ok) {
    $json = $resp.content | ConvertFrom-Json
    if ($null -ne $json.total) {
      return [long]$json.total
    }
  }

  # Fallback : pagination curseur complete
  $jqlOrdered = "$jql ORDER BY key ASC"
  $count = 0
  $nextToken = $null
  $isLast = $false
  $pageNum = 0

  do {
    $pageNum++
    $respPage = Invoke-JqlSearch -BaseUrl $BaseUrl -Headers $Headers -Jql $jqlOrdered -MaxResults 100 -Fields @("key") -NextPageToken $nextToken
    if (-not $respPage.ok) { break }
    $jsonPage = $respPage.content | ConvertFrom-Json
    if ($jsonPage.issues) { $count += ($jsonPage.issues | Measure-Object).Count }

    # Si total disponible dans la reponse (fallback GET v2)
    if ($null -ne $jsonPage.total) { return [long]$jsonPage.total }

    $isLast = $jsonPage.isLast
    $nextToken = $jsonPage.nextPageToken

    if ($pageNum % 20 -eq 0) {
      Log ("    Comptage Done {0} : {1} tickets (page {2})..." -f $ProjectKey, $count, $pageNum)
    }
    Start-Sleep -Milliseconds 100
  } while (-not $isLast -and $nextToken)

  return [long]$count
}

# ============================================================
# ECHANTILLONNAGE STRATIFIE PAR %DONE
# ============================================================

function Get-StratifiedSample {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$ProjectKey,
    [int]$SampleSize,
    [int]$TotalIssues,
    [int]$DoneCount,
    [int]$NotDoneCount
  )

  $issues = New-Object System.Collections.Generic.List[object]

  if ($TotalIssues -le 0) { return $issues }

  # Allocation proportionnelle
  $donePct = [double]$DoneCount / [double]$TotalIssues
  $sampleDone = [math]::Round($SampleSize * $donePct)
  $sampleNotDone = $SampleSize - $sampleDone

  # Garantir au moins 1 dans chaque strate si non vide
  if ($DoneCount -gt 0 -and $sampleDone -eq 0) { $sampleDone = 1; $sampleNotDone = $SampleSize - 1 }
  if ($NotDoneCount -gt 0 -and $sampleNotDone -eq 0) { $sampleNotDone = 1; $sampleDone = $SampleSize - 1 }

  # Cap aux tailles reelles
  if ($sampleDone -gt $DoneCount) { $sampleDone = $DoneCount; $sampleNotDone = $SampleSize - $sampleDone }
  if ($sampleNotDone -gt $NotDoneCount) { $sampleNotDone = $NotDoneCount; $sampleDone = $SampleSize - $sampleNotDone }

  # Fonction interne : paginer un stratum
  function Collect-Stratum {
    param([string]$Jql, [int]$Needed)
    $collected = New-Object System.Collections.Generic.List[object]
    if ($Needed -le 0) { return $collected }

    $nextToken = $null
    $pageNum = 0
    $maxPagesStratum = [math]::Ceiling($Needed / 100) + 2

    do {
      $pageNum++
      $resp = Invoke-JqlSearch -BaseUrl $BaseUrl -Headers $Headers -Jql $Jql -MaxResults 100 `
        -Fields @("summary", "description", "comment") -NextPageToken $nextToken
      if (-not $resp.ok) { break }
      $json = $resp.content | ConvertFrom-Json
      if ($json.issues) {
        foreach ($iss in $json.issues) {
          $collected.Add($iss) | Out-Null
          if ($collected.Count -ge $Needed) { break }
        }
      }
      if ($collected.Count -ge $Needed) { break }

      $isLast = $json.isLast
      $nextToken = $json.nextPageToken

      if ($null -ne $json.total -and -not $nextToken) {
        $isLast = ($collected.Count -ge [int]$json.total)
      }

      Start-Sleep -Milliseconds 150
    } while (-not $isLast -and $nextToken -and $pageNum -lt $maxPagesStratum)

    return $collected
  }

  # Strate Done
  if ($sampleDone -gt 0) {
    $jqlDone = "project = `"$ProjectKey`" AND statusCategory = Done ORDER BY key ASC"
    $doneIssues = Collect-Stratum -Jql $jqlDone -Needed $sampleDone
    foreach ($iss in $doneIssues) { $issues.Add($iss) | Out-Null }
  }

  # Strate NotDone
  if ($sampleNotDone -gt 0) {
    $jqlNotDone = "project = `"$ProjectKey`" AND statusCategory != Done ORDER BY key ASC"
    $notDoneIssues = Collect-Stratum -Jql $jqlNotDone -Needed $sampleNotDone
    foreach ($iss in $notDoneIssues) { $issues.Add($iss) | Out-Null }
  }

  return $issues
}

# ============================================================
# DETECTION DONNEES SENSIBLES
# ============================================================

$regexRules = @(
  @{ Name = "Tel FR mobile";     Pattern = '(?<!\d)(?:\+33\s?|0)(?:6|7)(?:[\s.\-]?\d{2}){4}(?!\d)'; Confidence = "Haute" }
  @{ Name = "Tel FR fixe";       Pattern = '(?<!\d)(?:\+33\s?|0)(?:1|2|3|4|5)(?:[\s.\-]?\d{2}){4}(?!\d)'; Confidence = "Moyenne" }
  @{ Name = "Tel international"; Pattern = '(?<!\d)\+(?:3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9])\s?\d(?:[\s.\-]?\d){7,12}(?!\d)'; Confidence = "Moyenne" }
  @{ Name = "IBAN FR";           Pattern = '(?<!\w)FR\s?\d{2}[\s.\-]?(?:\d{4}[\s.\-]?){5}\d{3}(?!\w)'; Confidence = "Haute" }
  @{ Name = "IBAN international"; Pattern = '(?<!\w)[A-Z]{2}\s?\d{2}[\s.\-]?[A-Z0-9]{4}(?:[\s.\-]?[A-Z0-9]{4}){2,7}(?:[\s.\-]?[A-Z0-9]{1,4})?(?!\w)'; Confidence = "Moyenne" }
  @{ Name = "NIR avec cle";     Pattern = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{2}(?!\d)'; Confidence = "Haute" }
  @{ Name = "NIR sans cle";     Pattern = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}(?!\d)'; Confidence = "Moyenne" }
  @{ Name = "Adresse voie";     Pattern = '(?i)(?<!\w)\d{1,4}[\s,]+(?:rue|avenue|boulevard|impasse|allee|place|chemin|route|passage|cours|square|residence|lotissement)\s+[A-Za-z\u00C0-\u017F\s\-]{3,40}(?!\w)'; Confidence = "Haute" }
  @{ Name = "Code postal ville"; Pattern = '(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5]|97[1-6])\d{3}\s+[A-Z\u00C0-\u017F][A-Za-z\u00C0-\u017F\s\-]{2,30}(?!\w)'; Confidence = "Moyenne" }
)

$exclusionPatterns = @(
  '^\d{4}-\d{2}-\d{2}'
  '^\d{2}/\d{2}/\d{4}'
  '^[A-Z]{2,10}-\d+'
  '^v?\d+\.\d+\.\d+'
  '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
  '^https?://'
)

function ConvertFrom-AdfNode($node) {
  if ($null -eq $node) { return "" }
  $text = ""
  if ($node.type -eq "text" -and $node.text) { $text += [string]$node.text + " " }
  if ($node.content) {
    foreach ($child in $node.content) { $text += ConvertFrom-AdfNode $child }
  }
  return $text
}

function Find-SensitiveDataInText {
  param([string]$Text, [string]$IssueKey, [string]$FieldName)
  if (-not $Text -or $Text.Length -lt 5) { return @() }
  $findings = @()
  foreach ($rule in $regexRules) {
    $rxMatches = [regex]::Matches($Text, $rule.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $rxMatches) {
      $matchValue = $m.Value.Trim()
      $excluded = $false
      foreach ($exPattern in $exclusionPatterns) {
        if ($matchValue -match $exPattern) { $excluded = $true; break }
      }
      if ($excluded) { continue }
      $findings += @{
        IssueKey   = $IssueKey
        Field      = $FieldName
        RuleName   = $rule.Name
        Confidence = $rule.Confidence
        MatchValue = $matchValue
      }
    }
  }
  return $findings
}

function Find-SensitiveDataInIssue {
  param($Issue)
  $issueKey = [string]$Issue.key
  $allFindings = @()

  # Summary
  if ($Issue.fields.summary) {
    $allFindings += Find-SensitiveDataInText -Text ([string]$Issue.fields.summary) -IssueKey $issueKey -FieldName "Summary"
  }

  # Description (ADF ou texte)
  if ($Issue.fields.description) {
    $descText = ""
    if ($Issue.fields.description.type -eq "doc") {
      $descText = ConvertFrom-AdfNode $Issue.fields.description
    } elseif ($Issue.fields.description -is [string]) {
      $descText = [string]$Issue.fields.description
    } else {
      try { $descText = $Issue.fields.description | ConvertTo-Json -Depth 10 } catch {}
    }
    $allFindings += Find-SensitiveDataInText -Text $descText -IssueKey $issueKey -FieldName "Description"
  }

  # Commentaires
  if ($Issue.fields.comment -and $Issue.fields.comment.comments) {
    $cIdx = 0
    foreach ($comment in $Issue.fields.comment.comments) {
      $cIdx++
      $commentText = ""
      if ($comment.body) {
        if ($comment.body.type -eq "doc") {
          $commentText = ConvertFrom-AdfNode $comment.body
        } elseif ($comment.body -is [string]) {
          $commentText = [string]$comment.body
        } else {
          try { $commentText = $comment.body | ConvertTo-Json -Depth 10 } catch {}
        }
      }
      $fieldLabel = "Commentaire {0}" -f $cIdx
      $allFindings += Find-SensitiveDataInText -Text $commentText -IssueKey $issueKey -FieldName $fieldLabel
    }
  }

  return $allFindings
}

# ============================================================
# STATISTIQUES (Cochran + Wilson IC)
# ============================================================

function Get-ZScore([double]$Confidence) {
  switch ([math]::Round($Confidence, 2)) {
    0.90 { return 1.645 }
    0.95 { return 1.96 }
    0.99 { return 2.576 }
    default { return 1.96 }
  }
}

function Get-CochranSampleSize {
  param([int]$Population, [double]$Margin, [double]$Confidence)
  $z = Get-ZScore $Confidence
  $p = 0.5
  $n0 = [math]::Ceiling(($z * $z * $p * (1 - $p)) / ($Margin * $Margin))
  if ($Population -gt 0) {
    $n = [math]::Ceiling([double]$n0 / (1.0 + ([double]($n0 - 1) / [double]$Population)))
  } else {
    $n = $n0
  }
  return $n
}

function Get-ConfidenceInterval {
  param([int]$SampleSize, [int]$Positives, [int]$Population, [double]$Confidence)
  if ($SampleSize -eq 0) { return @{ Lower = 0; Upper = 0; Estimate = 0 } }

  $z = Get-ZScore $Confidence
  $n = [double]$SampleSize
  $x = [double]$Positives
  $N = [double]$Population

  $pHat = $x / $n
  $denom = 1.0 + ($z * $z / $n)
  $center = ($pHat + ($z * $z) / (2.0 * $n)) / $denom
  $spread = ($z / $denom) * [math]::Sqrt(($pHat * (1.0 - $pHat) / $n) + ($z * $z / (4.0 * $n * $n)))

  $lower = [math]::Max(0.0, $center - $spread)
  $upper = [math]::Min(1.0, $center + $spread)

  # Correction population finie
  if ($N -gt 0 -and $n -lt $N) {
    $fpc = [math]::Sqrt(($N - $n) / ($N - 1.0))
    $adjustedSpread = $spread * $fpc
    $lower = [math]::Max(0.0, $center - $adjustedSpread)
    $upper = [math]::Min(1.0, $center + $adjustedSpread)
  }

  $estimate = [math]::Round($pHat * $N)
  $lowerCount = [math]::Round($lower * $N)
  $upperCount = [math]::Round($upper * $N)

  return @{ Lower = [int]$lowerCount; Upper = [int]$upperCount; Estimate = [int]$estimate }
}

# ============================================================
# PROGRAMME PRINCIPAL
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  SCAN RGPD - ESTIMATION STATISTIQUE" -ForegroundColor Cyan
Write-Host "  v1.13.2 - Comptage optimise + stratification" -ForegroundColor Cyan
Write-Host ("  Marge d'erreur : +/-{0:P1}" -f $MarginOfError) -ForegroundColor Cyan
Write-Host ("  Script dir : {0}" -f $ScriptDir) -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Credentials ---
$credPath = Join-Path $SecretsDir "site-admin.xml"
Log ("Chargement credentials : {0}" -f $credPath)
$site = Import-SiteCredentials -XmlPath $credPath
Log ("Site : {0}" -f $site.BaseUrl)

# --- Choix des projets ---
if (-not $Projects) {
  $userInput = Read-Host "  Cles des projets (ex: CPT,CNT,PACT ou * pour tous)"
  $Projects = $userInput.Trim()
}

$projectKeys = @()
if ($Projects -eq "*") {
  Log "Chargement de la liste de tous les projets..."
  $url = "{0}/rest/api/3/project?maxResults=500" -f $site.BaseUrl
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers
  if ($resp.ok) {
    $allProjects = $resp.content | ConvertFrom-Json
    $projectKeys = $allProjects | ForEach-Object { [string]$_.key }
    Log ("  {0} projets trouves" -f $projectKeys.Count)
  } else {
    Log "Erreur chargement liste projets : status=$($resp.status)" "ERROR"
    return
  }
} else {
  $projectKeys = $Projects.Split(",") | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ }
}

Log ("Projets a analyser : {0}" -f ($projectKeys -join ", "))
Log ("Marge erreur : +/-{0:P1} | Confiance : {1:P0}" -f $MarginOfError, $ConfidenceLevel)
Write-Host ""

# ============================================================
# BOUCLE SUR LES PROJETS
# ============================================================

$results = New-Object System.Collections.Generic.List[object]
$archivedCount = 0
$emptyCount = 0
$errorCount = 0
$globalTotalIssues = [long]0
$globalSampleSize = [long]0
$globalPositives = [long]0
$globalSampleDone = [long]0

foreach ($pk in $projectKeys) {

  $projectStart = Get-Date

  # --- Info projet : nom, categorie, lead ---
  $projUrl = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $pk
  $respProj = Invoke-ApiCall -Method "GET" -Url $projUrl -Headers $site.Headers

  $projectName = $pk
  $projectCategory = "(sans categorie)"
  $projectLead = "(sans lead)"

  if ($respProj.ok) {
    $projJson = $respProj.content | ConvertFrom-Json
    if ($projJson.name) { $projectName = [string]$projJson.name }
    if ($projJson.projectCategory -and $projJson.projectCategory.name) {
      $projectCategory = [string]$projJson.projectCategory.name
    }
    if ($projJson.lead -and $projJson.lead.displayName) {
      $projectLead = [string]$projJson.lead.displayName
    }
  } elseif ($respProj.status -eq 404) {
    Log ("  {0} : projet introuvable (404) - IGNORE" -f $pk) "WARN"
    $errorCount++
    continue
  } elseif ($respProj.status -eq 410) {
    Log ("  {0} : projet archive (410 Gone) - IGNORE" -f $pk)
    $archivedCount++
    continue
  }

  # --- Comptage total + derniere MAJ (UNE SEULE PASSE) ---
  $countInfo = Get-ProjectCountAndLastUpdated -BaseUrl $site.BaseUrl -Headers $site.Headers -ProjectKey $pk

  if (-not $countInfo.ok) {
    if ($countInfo.status -eq 410) {
      Log ("  {0} : projet archive (410 Gone) - IGNORE" -f $pk)
      $archivedCount++
    } else {
      Log ("  {0} : erreur API status={1}" -f $pk, $countInfo.status) "WARN"
      $errorCount++
    }
    Start-Sleep -Milliseconds 100
    continue
  }

  $totalIssues = [long]$countInfo.total
  $lastUpdated = $countInfo.lastUpdated

  if ($totalIssues -eq 0) {
    Log ("  {0} ({1}) : 0 tickets - IGNORE" -f $pk, $projectName) "WARN"
    $emptyCount++
    continue
  }

  # --- Comptage Done (une seule passe, NotDone = Total - Done) ---
  $doneCount = Get-DoneCount -BaseUrl $site.BaseUrl -Headers $site.Headers -ProjectKey $pk
  $notDoneCount = [long]$totalIssues - [long]$doneCount
  if ($notDoneCount -lt 0) { $notDoneCount = 0 }

  $pctDoneProject = [math]::Round(100.0 * [double]$doneCount / [double]$totalIssues, 1)

  # --- Taille echantillon ---
  $sampleSize = Get-CochranSampleSize -Population ([int]$totalIssues) -Margin $MarginOfError -Confidence $ConfidenceLevel
  $actualSampleSize = [Math]::Min($sampleSize, [int]$totalIssues)
  $isExhaustive = ($actualSampleSize -ge $totalIssues)

  $pctEchantillon = [math]::Round(100.0 * [double]$actualSampleSize / [double]$totalIssues, 1)

  $scanLabel = if ($isExhaustive) { "EXHAUSTIF" } else { "echantillon $actualSampleSize/$totalIssues ({0}%)" -f $pctEchantillon }
  Log ("  {0} ({1}) | Cat: {2} | Lead: {3} | MAJ: {4} | {5} tickets ({6}% Done) | {7}" -f `
    $pk, $projectName, $projectCategory, $projectLead, $lastUpdated, $totalIssues, $pctDoneProject, $scanLabel)

  # --- Echantillonnage stratifie et scan ---
  Write-Progress -Activity ("Scan {0}" -f $pk) -Status "Recuperation des tickets (stratifie)..." -PercentComplete 0

  $sampleIssues = Get-StratifiedSample -BaseUrl $site.BaseUrl -Headers $site.Headers `
    -ProjectKey $pk -SampleSize $actualSampleSize -TotalIssues ([int]$totalIssues) `
    -DoneCount ([int]$doneCount) -NotDoneCount ([int]$notDoneCount)

  $actualSampleSize = ($sampleIssues | Measure-Object).Count
  $pctEchantillon = if ($totalIssues -gt 0) { [math]::Round(100.0 * [double]$actualSampleSize / [double]$totalIssues, 1) } else { 0.0 }

  # %Done dans l echantillon
  $sampleDoneCount = [math]::Min(
    [math]::Round([double]$actualSampleSize * [double]$doneCount / [math]::Max(1, [double]$totalIssues)),
    $doneCount
  )
  $pctDoneSample = if ($actualSampleSize -gt 0) { [math]::Round(100.0 * [double]$sampleDoneCount / [double]$actualSampleSize, 1) } else { 0.0 }

  # Verification ecart
  $ecart = [math]::Abs($pctDoneProject - $pctDoneSample)
  if ($ecart -gt 5.0) {
    Log ("    WARNING: ecart %Done projet ({0}%) vs echantillon ({1}%) = {2}%" -f $pctDoneProject, $pctDoneSample, $ecart) "WARN"
  }

  # --- Scan des tickets ---
  $positives = 0
  $idx = 0

  foreach ($issue in $sampleIssues) {
    $idx++
    $findings = Find-SensitiveDataInIssue -Issue $issue
    if ($findings.Count -gt 0) { $positives++ }
    if ($idx % 50 -eq 0) {
      Write-Progress -Activity ("Scan {0}" -f $pk) `
        -Status ("{0}/{1} scannes, {2} positifs" -f $idx, $actualSampleSize, $positives) `
        -PercentComplete ([int](100 * $idx / [math]::Max(1, $actualSampleSize)))
    }
  }

  Write-Progress -Activity ("Scan {0}" -f $pk) -Completed

  # --- % sensibles ---
  $pctSensibles = if ($actualSampleSize -gt 0) { [math]::Round(100.0 * [double]$positives / [double]$actualSampleSize, 2) } else { 0.0 }

  # --- IC ---
  $ci = Get-ConfidenceInterval -SampleSize $actualSampleSize -Positives $positives `
    -Population ([int]$totalIssues) -Confidence $ConfidenceLevel

  # --- Temps ecoule ---
  $elapsed = ((Get-Date) - $projectStart).TotalSeconds

  $results.Add(@{
    ProjectKey      = $pk
    ProjectName     = $projectName
    Category        = $projectCategory
    Lead            = $projectLead
    LastUpdated     = $lastUpdated
    TotalIssues     = [long]$totalIssues
    DoneCount       = [long]$doneCount
    NotDoneCount    = [long]$notDoneCount
    PctDoneProject  = $pctDoneProject
    SampleSize      = $actualSampleSize
    PctEchantillon  = $pctEchantillon
    SampleDone      = $sampleDoneCount
    PctDoneSample   = $pctDoneSample
    Positives       = $positives
    PctSensibles    = $pctSensibles
    Estimate        = $ci.Estimate
    LowerBound      = $ci.Lower
    UpperBound      = $ci.Upper
    IsExhaustive    = $isExhaustive
  }) | Out-Null

  $globalTotalIssues += $totalIssues
  $globalSampleSize += $actualSampleSize
  $globalPositives += $positives
  $globalSampleDone += $sampleDoneCount

  Log ("    => {0} positifs / {1} scannes ({2}%) -> estimation {3} [{4} - {5}] ({6:N0}s)" -f `
    $positives, $actualSampleSize, $pctSensibles, $ci.Estimate, $ci.Lower, $ci.Upper, $elapsed)

  Start-Sleep -Milliseconds 200
}

# ============================================================
# ESTIMATION GLOBALE
# ============================================================

$globalCI = Get-ConfidenceInterval -SampleSize ([int]$globalSampleSize) -Positives ([int]$globalPositives) `
  -Population ([int]$globalTotalIssues) -Confidence $ConfidenceLevel

$globalPctDone = if ($globalSampleSize -gt 0) {
  [math]::Round(100.0 * [double]$globalSampleDone / [double]$globalSampleSize, 1)
} else { 0.0 }

$globalPctEchantillon = if ($globalTotalIssues -gt 0) {
  [math]::Round(100.0 * [double]$globalSampleSize / [double]$globalTotalIssues, 1)
} else { 0.0 }

$globalPctSensibles = if ($globalSampleSize -gt 0) {
  [math]::Round(100.0 * [double]$globalPositives / [double]$globalSampleSize, 2)
} else { 0.0 }

# ============================================================
# TRI PAR CATEGORIE + CLE PROJET
# ============================================================

$sortedResults = $results | Sort-Object { $_.Category }, { $_.ProjectKey }

# ============================================================
# EXPORT CSV
# ============================================================

Log "=== Export CSV ==="

$csvHeader = '"Categorie";"Projet";"Nom";"Lead";"Derniere MAJ";"Total";"Done";"NotDone";"%Done Projet";"Echantillon";"%Echantillon";"Ech. Done";"%Done Ech.";"Sensibles";"%Sensibles";"Estimation";"IC Bas";"IC Haut";"Exhaustif"'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($csvFile, "$csvHeader`r`n", $utf8Bom)

foreach ($r in $sortedResults) {
  $line = '"{0}";"{1}";"{2}";"{3}";"{4}";{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};"{18}"' -f `
    ($r.Category -replace '"','""'),
    $r.ProjectKey,
    ($r.ProjectName -replace '"','""'),
    ($r.Lead -replace '"','""'),
    $r.LastUpdated,
    $r.TotalIssues,
    $r.DoneCount,
    $r.NotDoneCount,
    $r.PctDoneProject,
    $r.SampleSize,
    $r.PctEchantillon,
    $r.SampleDone,
    $r.PctDoneSample,
    $r.Positives,
    $r.PctSensibles,
    $r.Estimate,
    $r.LowerBound,
    $r.UpperBound,
    $(if ($r.IsExhaustive) { "Oui" } else { "Non" })
  Add-CsvLine $line
}

# Ligne globale
$globalLine = '"GLOBAL";"";"TOUS LES PROJETS";"";"";{0};"";"";"";{1};{2};{3};{4};{5};{6};{7};{8};{9};""' -f `
  $globalTotalIssues, $globalSampleSize, $globalPctEchantillon, $globalSampleDone, $globalPctDone,
  $globalPositives, $globalPctSensibles, $globalCI.Estimate, $globalCI.Lower, $globalCI.Upper
Add-CsvLine $globalLine

Log ("CSV -> {0}" -f $csvFile)

# ============================================================
# AFFICHAGE DES RESULTATS
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESULTATS (tries par Categorie)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$currentCategory = ""
foreach ($r in $sortedResults) {
  if ($r.Category -ne $currentCategory) {
    $currentCategory = $r.Category
    Write-Host ""
    Write-Host ("  --- {0} ---" -f $currentCategory) -ForegroundColor Yellow
  }
  $exLabel = if ($r.IsExhaustive) { "[EXHAUSTIF]" } else { "" }
  Write-Host ("    {0,-8} | {1,-30} | Lead: {2,-20} | MAJ: {3}" -f $r.ProjectKey, $r.ProjectName, $r.Lead, $r.LastUpdated) -ForegroundColor White
  Write-Host ("             Total: {0,6} | Echant.: {1,5} ({2,5}%) | Sensibles: {3,4} ({4}%) | Est.: {5,5} [{6} - {7}] {8}" -f `
    $r.TotalIssues, $r.SampleSize, $r.PctEchantillon, $r.Positives, $r.PctSensibles, $r.Estimate, $r.LowerBound, $r.UpperBound, $exLabel) -ForegroundColor DarkGray
  Write-Host ("             %Done Projet: {0,5}% | %Done Echant.: {1,5}%" -f $r.PctDoneProject, $r.PctDoneSample) -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  ========================================" -ForegroundColor Green
Write-Host ("  GLOBAL : {0} projets analyses | {1} tickets" -f $results.Count, $globalTotalIssues) -ForegroundColor Green
Write-Host ("           Echantillon total : {0} tickets ({1}% du total)" -f $globalSampleSize, $globalPctEchantillon) -ForegroundColor Green
Write-Host ("           %Done echantillon : {0}%" -f $globalPctDone) -ForegroundColor Green
Write-Host ("           Sensibles trouves : {0} / {1} ({2}%)" -f $globalPositives, $globalSampleSize, $globalPctSensibles) -ForegroundColor Green
Write-Host ("           Estimation : {0} tickets sensibles [{1} - {2}]" -f $globalCI.Estimate, $globalCI.Lower, $globalCI.Upper) -ForegroundColor Green
Write-Host ("           Archives ignores : {0} | Vides ignores : {1}" -f $archivedCount, $emptyCount) -ForegroundColor Green
Write-Host "  ========================================" -ForegroundColor Green

# ============================================================
# RESUME
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Projets analyses          : {0}" -f $results.Count)
Write-Host ("  Projets archives ignores  : {0}" -f $archivedCount)
Write-Host ("  Projets vides ignores     : {0}" -f $emptyCount)
Write-Host ("  Projets en erreur         : {0}" -f $errorCount)
Write-Host ("  Tickets totaux (pop.)     : {0}" -f $globalTotalIssues)
Write-Host ("  Tickets scannes (echant.) : {0} ({1}% du total)" -f $globalSampleSize, $globalPctEchantillon)
Write-Host ("  Tickets sensibles trouves : {0} ({1}%)" -f $globalPositives, $globalPctSensibles)
Write-Host ("  Estimation globale        : {0} [{1} - {2}]" -f $globalCI.Estimate, $globalCI.Lower, $globalCI.Upper)
Write-Host ("  Marge erreur              : +/-{0:P1}" -f $MarginOfError)
Write-Host ("  Confiance                 : {0:P0}" -f $ConfidenceLevel)
Write-Host ("  CSV                       : {0}" -f $csvFile)
Write-Host ("  LOG                       : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."
