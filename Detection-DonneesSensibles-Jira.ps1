<#
.SYNOPSIS
  Detection-DonneesSensibles-Jira.ps1
  Detecte les donnees personnelles sensibles dans les tickets Jira
  d'un projet donne : telephones, IBAN, numeros de Securite Sociale,
  adresses postales de personnes physiques.

.DESCRIPTION
  Ce script :
  1. Demande la cle du projet Jira en console
  2. Parcourt tous les tickets du projet (summary, description, commentaires, champs custom texte)
  3. Applique des regex pour detecter :
     - Numeros de telephone (FR fixe, mobile, international)
     - IBAN (FR et internationaux, avec validation longueur)
     - Numeros de Securite Sociale (avec validation cle)
     - Adresses postales (code postal + ville, avec mots-cles)
  4. Genere un rapport CSV et un log

  ANTI-FAUX POSITIFS :
    - Validation de la cle de controle pour les NIR (Secu)
    - Validation de la longueur IBAN par pays
    - Exclusion des patterns connus (JIRA keys, dates, IP, URLs)
    - Contexte : extraction de 50 caracteres autour du match
    - Score de confiance (Haute / Moyenne / Basse)

  API :
    POST /rest/api/3/search/jql (nouvel endpoint Jira Cloud)
    Pagination par nextPageToken dans le body JSON

  CREDENTIALS :
    secrets\site-admin.xml => Jiradot (SiteUrl + Email + API Token)

  FICHIERS GENERES (dans exports\) :
    DonneesSensibles_{projet}_{ts}.csv   => Rapport detaille
    DonneesSensibles_{projet}_{ts}.log   => Journal d execution

.NOTES
  Auteur         : Frederic GUEDJ
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Detection-DonneesSensibles-Jira.ps1
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

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  if ($script:logFile) { Add-Content -Path $script:logFile -Value $line -Encoding UTF8 }
}

function Write-CsvHeader([string]$line) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($script:csvFile, "$line`r`n", $utf8Bom)
}

function Write-CsvLine([string]$line) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::AppendAllText($script:csvFile, "$line`r`n", $utf8Bom)
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
# REGEX ET DETECTION
# ============================================================

$regexTelFR = @(
  @{
    Name        = "Tel FR mobile"
    Pattern     = '(?<!\d)(?:\+33\s?|0)(?:6|7)(?:[\s.\-]?\d{2}){4}(?!\d)'
    Confidence  = "Haute"
    Description = "Telephone mobile FR (06/07)"
  },
  @{
    Name        = "Tel FR fixe"
    Pattern     = '(?<!\d)(?:\+33\s?|0)(?:1|2|3|4|5)(?:[\s.\-]?\d{2}){4}(?!\d)'
    Confidence  = "Moyenne"
    Description = "Telephone fixe FR (01-05)"
  },
  @{
    Name        = "Tel FR services"
    Pattern     = '(?<!\d)(?:\+33\s?|0)(?:8|9)(?:[\s.\-]?\d{2}){4}(?!\d)'
    Confidence  = "Basse"
    Description = "Telephone services FR (08/09) - peut etre un numero technique"
  },
  @{
    Name        = "Tel international"
    Pattern     = '(?<!\d)\+(?:3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9])\s?\d(?:[\s.\-]?\d){7,12}(?!\d)'
    Confidence  = "Moyenne"
    Description = "Telephone international (+XX...)"
  }
)

$regexIBAN = @(
  @{
    Name        = "IBAN FR"
    Pattern     = '(?<!\w)FR\s?\d{2}[\s.\-]?(?:\d{4}[\s.\-]?){5}\d{3}(?!\w)'
    Confidence  = "Haute"
    Description = "IBAN francais (FR76 + 23 chiffres)"
  },
  @{
    Name        = "IBAN international"
    Pattern     = '(?<!\w)[A-Z]{2}\s?\d{2}[\s.\-]?[A-Z0-9]{4}(?:[\s.\-]?[A-Z0-9]{4}){2,7}(?:[\s.\-]?[A-Z0-9]{1,4})?(?!\w)'
    Confidence  = "Moyenne"
    Description = "IBAN international (XX00 + alphanum)"
  }
)

$regexNIR = @(
  @{
    Name        = "NIR (Secu) avec cle"
    Pattern     = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{2}(?!\d)'
    Confidence  = "Haute"
    Description = "Numero de Securite Sociale avec cle de controle"
  },
  @{
    Name        = "NIR (Secu) sans cle"
    Pattern     = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}(?!\d)'
    Confidence  = "Moyenne"
    Description = "Numero de Securite Sociale sans cle (13 chiffres)"
  }
)

$regexAdresse = @(
  @{
    Name        = "Adresse avec numero et voie"
    Pattern     = '(?<!\w)\d{1,4}[\s,]+(?:rue|avenue|boulevard|impasse|allee|place|chemin|route|passage|cours|square|residence|lotissement|hameau|lieu[\s-]?dit|av\.|bd\.|bld\.|r\.|pl\.)\s+[A-Za-z\u00C0-\u017F\s\-'']{3,40}(?!\w)'
    Confidence  = "Haute"
    Description = "Adresse avec numero et type de voie"
  },
  @{
    Name        = "Code postal + ville"
    Pattern     = '(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5]|97[1-6]|98[4-9])\d{3}\s+[A-Z\u00C0-\u017F][A-Za-z\u00C0-\u017F\s\-'']{2,30}(?!\w)'
    Confidence  = "Moyenne"
    Description = "Code postal FR + nom de ville"
  }
)

$allRules = @()
$allRules += $regexTelFR
$allRules += $regexIBAN
$allRules += $regexNIR
$allRules += $regexAdresse

$exclusionPatterns = @(
  '^\d{4}-\d{2}-\d{2}',
  '^\d{2}/\d{2}/\d{4}',
  '^[A-Z]{2,10}-\d+',
  '^v?\d+\.\d+\.\d+',
  '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}',
  '^https?://'
)

# ============================================================
# FONCTIONS DE DETECTION
# ============================================================

function Test-Exclusion([string]$matchValue) {
  foreach ($exPattern in $exclusionPatterns) {
    if ($matchValue -match $exPattern) { return $true }
  }
  return $false
}

function Test-NIRChecksum([string]$nirRaw) {
  $nir = $nirRaw -replace '[\s.\-]', ''
  $nirCalc = $nir
  if ($nirCalc -match '2A') { $nirCalc = $nirCalc -replace '2A', '19' }
  if ($nirCalc -match '2B') { $nirCalc = $nirCalc -replace '2B', '18' }
  if ($nirCalc.Length -eq 15 -and $nirCalc -match '^\d{15}$') {
    try {
      $base = [long]$nirCalc.Substring(0, 13)
      $cle  = [int]$nirCalc.Substring(13, 2)
      $cleCalc = 97 - ($base % 97)
      return ($cle -eq $cleCalc)
    } catch { return $false }
  }
  return $true
}

function Get-Context([string]$text, [int]$index, [int]$length, [int]$contextSize=50) {
  $start = [Math]::Max(0, $index - $contextSize)
  $end   = [Math]::Min($text.Length, $index + $length + $contextSize)
  $ctx   = $text.Substring($start, $end - $start)
  $ctx   = $ctx -replace '[\r\n]+', ' '
  $ctx   = $ctx -replace '\s+', ' '
  if ($start -gt 0) { $ctx = "..." + $ctx }
  if ($end -lt $text.Length) { $ctx = $ctx + "..." }
  return $ctx.Trim()
}

function Extract-TextFromADF($node) {
  if ($null -eq $node) { return "" }
  $text = ""
  if ($node.type -eq "text" -and $node.text) {
    $text += [string]$node.text + " "
  }
  if ($node.content) {
    foreach ($child in $node.content) {
      $text += Extract-TextFromADF $child
    }
  }
  return $text
}

function Scan-Text([string]$text, [string]$issueKey, [string]$fieldName) {
  if (-not $text -or $text.Length -lt 5) { return @() }
  $findings = @()

  foreach ($rule in $allRules) {
    $matches = [regex]::Matches($text, $rule.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    foreach ($m in $matches) {
      $matchValue = $m.Value.Trim()

      if (Test-Exclusion $matchValue) { continue }

      $confidence = $rule.Confidence
      if ($rule.Name -like "NIR*") {
        if (-not (Test-NIRChecksum $matchValue)) {
          $confidence = "Basse"
        }
      }

      if ($rule.Name -eq "IBAN FR") {
        $ibanClean = $matchValue -replace '[\s.\-]', ''
        if ($ibanClean.Length -ne 27) { $confidence = "Basse" }
      }

      $context = Get-Context $text $m.Index $m.Length

      $findings += @{
        IssueKey   = $issueKey
        Field      = $fieldName
        RuleName   = $rule.Name
        Category   = if ($rule.Name -like "Tel*") { "Telephone" }
                     elseif ($rule.Name -like "IBAN*") { "IBAN" }
                     elseif ($rule.Name -like "NIR*") { "Securite Sociale" }
                     elseif ($rule.Name -like "Adresse*" -or $rule.Name -like "Code*") { "Adresse" }
                     else { "Autre" }
        MatchValue = $matchValue
        Confidence = $confidence
        Context    = $context
      }
    }
  }

  return $findings
}

# ============================================================
# DIALOGUE : CHOIX DU PROJET
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  DETECTION DONNEES SENSIBLES DANS JIRA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Types detectes :" -ForegroundColor White
Write-Host "    - Numeros de telephone (FR, international)" -ForegroundColor DarkGray
Write-Host "    - IBAN (FR, international)" -ForegroundColor DarkGray
Write-Host "    - Numeros de Securite Sociale (NIR)" -ForegroundColor DarkGray
Write-Host "    - Adresses postales" -ForegroundColor DarkGray
Write-Host ""

$projectKey = ""
while (-not $projectKey) {
  $projectKey = (Read-Host "  Cle du projet Jira (ex: PACT, DOCJ)").Trim().ToUpper()
}

$logFile = Join-Path $ExportsDir ("DonneesSensibles_{0}_{1}.log" -f $projectKey, $ts)
$csvFile = Join-Path $ExportsDir ("DonneesSensibles_{0}_{1}.csv" -f $projectKey, $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

Log "================================================================"
Log "  DETECTION DONNEES SENSIBLES"
Log ("  Projet : {0}" -f $projectKey)
Log "================================================================"

# --- Credentials ---
$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
Log ("  Site : {0}" -f $site.BaseUrl)

# ============================================================
# ETAPE 1 : VERIFICATION DU PROJET
# ============================================================

Log "=== ETAPE 1 : Verification du projet ==="

$url = "{0}/rest/api/3/project/{1}" -f $site.BaseUrl, $projectKey
$resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

if (-not $resp.ok) {
  Log ("  Projet {0} non trouve ou inaccessible : status={1}" -f $projectKey, $resp.status) "ERROR"
  Write-Host ("  ERREUR : projet {0} non trouve (status={1})" -f $projectKey, $resp.status) -ForegroundColor Red
  return
}

$projectJson = $resp.content | ConvertFrom-Json
$projectName = [string]$projectJson.name
Log ("  Projet : {0} ({1})" -f $projectName, $projectKey)
Write-Host ("  => Projet : {0} ({1})" -f $projectName, $projectKey) -ForegroundColor Green

# Pre-check : verifier l'acces aux tickets
$preCheckBody = @{
  jql        = "project = $projectKey"
  maxResults = 1
  fields     = @("key")
} | ConvertTo-Json -Depth 5 -Compress

$preCheckUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl
$preCheckResp = Invoke-ApiCall -Method "POST" -Url $preCheckUrl -Headers $site.Headers -Body $preCheckBody

if (-not $preCheckResp.ok) {
  Log ("  Erreur acces tickets : status={0}" -f $preCheckResp.status) "ERROR"
  Write-Host ("  ERREUR : impossible d'acceder aux tickets (status={0})" -f $preCheckResp.status) -ForegroundColor Red
  return
}

$preCheckJson = $preCheckResp.content | ConvertFrom-Json
$preCheckIssues = ($preCheckJson.issues | Measure-Object).Count

if ($preCheckIssues -eq 0) {
  Write-Host "  Aucun ticket dans ce projet." -ForegroundColor Yellow
  return
}

Write-Host "  Acces aux tickets OK, demarrage du scan..." -ForegroundColor Green
Write-Host ""

# ============================================================
# ETAPE 2 : SCAN DES TICKETS
#
# Nouvel endpoint POST /rest/api/3/search/jql :
#   - PAS de startAt (ni body ni query string)
#   - Pagination via nextPageToken DANS LE BODY
#   - Le body complet (jql, fields, maxResults) est requis a chaque page
#   - Reponse : issues[], nextPageToken, isLast
# ============================================================

Log "=== ETAPE 2 : Scan des tickets ==="

$allFindings = New-Object System.Collections.Generic.List[object]
$cScanned = 0; $cWithFindings = 0; $cPages = 0
$startTime = Get-Date
$nextPageToken = $null
$hasMore = $true
$pageSize = 50
$seenKeys = New-Object System.Collections.Generic.HashSet[string]

$searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl

while ($hasMore) {
  $cPages++

  # Construire le body — toujours complet, avec nextPageToken si disponible
  $bodyObj = @{
    jql        = "project = $projectKey ORDER BY key ASC"
    maxResults = $pageSize
    fields     = @("key", "summary", "description", "comment")
  }
  if ($nextPageToken) {
    $bodyObj["nextPageToken"] = $nextPageToken
  }
  $searchBody = $bodyObj | ConvertTo-Json -Depth 5 -Compress

  $resp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $searchBody

  if (-not $resp.ok) {
    Log ("  Erreur search page {0} : status={1} {2}" -f $cPages, $resp.status, $resp.error) "ERROR"
    break
  }

  $json = $resp.content | ConvertFrom-Json
  $issues = $json.issues
  $issueCount = ($issues | Measure-Object).Count

  if ($issueCount -eq 0) {
    $hasMore = $false
    break
  }

  # Detection de boucle
  $firstKey = [string]$issues[0].key
  if ($seenKeys.Contains($firstKey)) {
    Log ("  BOUCLE DETECTEE : {0} deja vu a la page {1}, arret force" -f $firstKey, $cPages) "WARN"
    $hasMore = $false
    break
  }

  foreach ($issue in $issues) {
    $issueKey = [string]$issue.key
    [void]$seenKeys.Add($issueKey)
    $cScanned++
    $issueFindings = @()

    # --- Summary ---
    if ($issue.fields.summary) {
      $summaryText = [string]$issue.fields.summary
      $issueFindings += Scan-Text $summaryText $issueKey "Summary"
    }

    # --- Description (ADF ou texte) ---
    if ($issue.fields.description) {
      $descText = ""
      if ($issue.fields.description.type -eq "doc") {
        $descText = Extract-TextFromADF $issue.fields.description
      } elseif ($issue.fields.description -is [string]) {
        $descText = [string]$issue.fields.description
      } else {
        try { $descText = $issue.fields.description | ConvertTo-Json -Depth 10 } catch {}
      }
      $issueFindings += Scan-Text $descText $issueKey "Description"
    }

    # --- Commentaires ---
    if ($issue.fields.comment -and $issue.fields.comment.comments) {
      $cIdx = 0
      foreach ($comment in $issue.fields.comment.comments) {
        $cIdx++
        $commentText = ""
        if ($comment.body) {
          if ($comment.body.type -eq "doc") {
            $commentText = Extract-TextFromADF $comment.body
          } elseif ($comment.body -is [string]) {
            $commentText = [string]$comment.body
          } else {
            try { $commentText = $comment.body | ConvertTo-Json -Depth 10 } catch {}
          }
        }
        $authorName = ""
        if ($comment.author -and $comment.author.displayName) {
          $authorName = [string]$comment.author.displayName
        }
        $fieldLabel = "Commentaire #{0} ({1})" -f $cIdx, $authorName
        $issueFindings += Scan-Text $commentText $issueKey $fieldLabel
      }
    }

    # --- Champs custom texte ---
    foreach ($prop in $issue.fields.PSObject.Properties) {
      if ($prop.Name -like "customfield_*" -and $prop.Value) {
        $cfText = ""
        if ($prop.Value -is [string]) {
          $cfText = [string]$prop.Value
        } elseif ($prop.Value.type -eq "doc") {
          $cfText = Extract-TextFromADF $prop.Value
        }
        if ($cfText.Length -gt 5) {
          $cfFindings = Scan-Text $cfText $issueKey $prop.Name
          $issueFindings += $cfFindings
        }
      }
    }

    if ($issueFindings.Count -gt 0) {
      $cWithFindings++
      foreach ($f in $issueFindings) {
        $allFindings.Add($f) | Out-Null
      }
    }
  }

  # Progression
  $elapsed = (Get-Date) - $startTime
  $rate = if ($elapsed.TotalSeconds -gt 0) { [Math]::Round($cScanned / $elapsed.TotalSeconds, 0) } else { 0 }
  $lastKey = [string]$issues[$issueCount - 1].key
  Write-Progress -Activity "Scan $projectKey" `
    -Status ("{0} scannes ({1}), {2} trouvailles ({3} t/s)" -f $cScanned, $lastKey, $allFindings.Count, $rate) `
    -PercentComplete (-1)

  Log ("  Page {0} : {1} tickets [{2}..{3}], cumul {4}, trouvailles {5}" -f $cPages, $issueCount, $firstKey, $lastKey, $cScanned, $allFindings.Count)

  # Pagination : nextPageToken dans la reponse
  if ($json.isLast -eq $true) {
    $hasMore = $false
  } elseif ($json.nextPageToken) {
    $nextPageToken = [string]$json.nextPageToken
  } else {
    $hasMore = $false
  }

  Start-Sleep -Milliseconds 200
}

Write-Progress -Activity "Scan" -Completed

Log ("  Tickets scannes       : {0}" -f $cScanned)
Log ("  Tickets avec donnees  : {0}" -f $cWithFindings)
Log ("  Trouvailles totales   : {0}" -f $allFindings.Count)
Log ("  Pages API             : {0}" -f $cPages)

# ============================================================
# ETAPE 3 : EXPORT CSV
# ============================================================

Log "=== ETAPE 3 : Export CSV ==="

$csvColumns = @("IssueKey","IssueUrl","Field","Category","RuleName","Confidence","MatchValue","Context")
$csvHeader = ($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
Write-CsvHeader $csvHeader

foreach ($f in ($allFindings | Sort-Object { $_.IssueKey })) {
  $issueUrl = "{0}/browse/{1}" -f $site.BaseUrl, $f.IssueKey
  $row = [ordered]@{
    IssueKey   = $f.IssueKey
    IssueUrl   = $issueUrl
    Field      = $f.Field
    Category   = $f.Category
    RuleName   = $f.RuleName
    Confidence = $f.Confidence
    MatchValue = $f.MatchValue
    Context    = $f.Context
  }
  $line = ($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";"
  Write-CsvLine $line
}

Log "  CSV -> $csvFile"

# ============================================================
# ETAPE 4 : AFFICHAGE DES RESULTATS
# ============================================================

$cTotal = ($allFindings | Measure-Object).Count
$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESULTATS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($cTotal -eq 0) {
  Write-Host "  Aucune donnee sensible detectee." -ForegroundColor Green
} else {
  # Stats par categorie
  $categories = $allFindings | Group-Object { $_.Category }
  foreach ($cat in ($categories | Sort-Object Name)) {
    $catCount = ($cat.Group | Measure-Object).Count
    $color = switch ($cat.Name) {
      "Telephone"        { "Yellow" }
      "IBAN"             { "Red" }
      "Securite Sociale" { "Red" }
      "Adresse"          { "DarkYellow" }
      default            { "White" }
    }
    Write-Host ("  {0,-20} : {1}" -f $cat.Name, $catCount) -ForegroundColor $color
  }

  Write-Host ""

  # Stats par confiance
  $confGroups = $allFindings | Group-Object { $_.Confidence }
  foreach ($cg in ($confGroups | Sort-Object Name)) {
    $cgCount = ($cg.Group | Measure-Object).Count
    $color = switch ($cg.Name) {
      "Haute"   { "Red" }
      "Moyenne" { "Yellow" }
      "Basse"   { "DarkGray" }
      default   { "White" }
    }
    Write-Host ("  Confiance {0,-10} : {1}" -f $cg.Name, $cgCount) -ForegroundColor $color
  }

  Write-Host ""

  # Top 10 tickets
  $topIssues = $allFindings | Group-Object { $_.IssueKey } | Sort-Object Count -Descending | Select-Object -First 10
  Write-Host "  Top 10 tickets :" -ForegroundColor White
  foreach ($ti in $topIssues) {
    $tiCount = ($ti.Group | Measure-Object).Count
    Write-Host ("    {0,-15} : {1} trouvailles" -f $ti.Name, $tiCount) -ForegroundColor Yellow
  }
}

# ============================================================
# RESUME
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Projet                : {0} ({1})" -f $projectName, $projectKey)
Write-Host ("  Tickets scannes       : {0}" -f $cScanned)
Write-Host ("  Pages API             : {0}" -f $cPages)
Write-Host ("  Tickets avec donnees  : {0}" -f $cWithFindings) -ForegroundColor $(if ($cWithFindings -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Trouvailles totales   : {0}" -f $cTotal) -ForegroundColor $(if ($cTotal -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Duree                 : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  CSV                   : {0}" -f $csvFile)
Write-Host ("  LOG                   : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."