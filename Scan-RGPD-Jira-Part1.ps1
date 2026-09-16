<#
.SYNOPSIS
  Scan-RGPD-Jira-Part1.ps1
  Detection en masse de donnees personnelles sensibles dans les tickets Jira.
  Partie 1 : Scan des champs texte (Summary, Description, Commentaires, Custom fields).

.DESCRIPTION
  Ce script :
  1. Recupere les categories de projets Jira du site
  2. Propose un menu de selection par categorie (ex: "software", "business", "service_desk")
  3. Scanne TOUS les projets de la categorie choisie
  4. Detecte : telephones, IBAN, numeros de Securite Sociale, adresses postales
  5. Exclut les adresses d'entreprise connues (Harmonie Mutuelle, Mutex, VYV, Chatillon, etc.)
  6. Genere un CSV (UTF-8 BOM, separateur ;) et un log

  ANTI-FAUX POSITIFS :
    - Validation cle de controle NIR
    - Validation longueur IBAN FR (27 car.) et IBAN international (15-34 car.)
    - IBAN international : regex durci (min 4 blocs, pas d'espace apres code pays)
    - Exclusion patterns connus (cles Jira, dates, IP, URLs, versions)
    - Exclusion adresses d'entreprise (Harmonie Mutuelle, Mutex, VYV, MGEN, MNT, Chatillon)
    - Contexte 50 caracteres autour du match
    - Score de confiance (Haute / Moyenne / Basse)

  API :
    GET  /rest/api/3/project/search  => liste des projets par categorie
    POST /rest/api/3/search/jql      => tickets (pagination nextPageToken)

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 2.3 - Partie 1 (sans PJ)
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Scan-RGPD-Jira-Part1.ps1
  .\Scan-RGPD-Jira-Part1.ps1 -CategoryFilter "software"
  .\Scan-RGPD-Jira-Part1.ps1 -CategoryFilter "software" -PageSize 100
#>

[CmdletBinding()]
param(
  [string] $CategoryFilter = "",      # Vide = menu interactif
  [int]    $PageSize       = 100,     # Tickets par page API (max 100)
  [int]    $MaxRetries     = 5,
  [int]    $ThrottleMs     = 200      # Pause entre pages API
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ============================================================
# CONSTANTES
# ============================================================

$VERSION = "2.3-Part1"

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$script:logFile = $null
$script:csvFile = $null

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level = "INFO") {
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

function Escape-CsvField([string]$value) {
  return '"{0}"' -f ($value -replace '"', '""')
}

# ============================================================
# RESEAU & PROXY
# ============================================================

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# ============================================================
# HTTP HELPER (avec retry exponentiel)
# ============================================================

function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)
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

      # Forcer UTF-8
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
        Log ("  Retry {0} status={1} dans {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
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
    @{ SiteUrl = $inputUrl; Email = $adminEmail; ApiTokenSecureString = $apiTokenSecure } |
      Export-Clixml -Path $CredFile
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
# REGEX DE DETECTION
# ============================================================

$regexTelFR = @(
  @{ Name = "Tel FR mobile";     Pattern = '(?<!\d)(?:\+33\s?|0)(?:6|7)(?:[\s.\-]?\d{2}){4}(?!\d)';  Confidence = "Haute";   Category = "Telephone" },
  @{ Name = "Tel FR fixe";       Pattern = '(?<!\d)(?:\+33\s?|0)(?:1|2|3|4|5)(?:[\s.\-]?\d{2}){4}(?!\d)'; Confidence = "Moyenne"; Category = "Telephone" },
  @{ Name = "Tel FR services";   Pattern = '(?<!\d)(?:\+33\s?|0)(?:8|9)(?:[\s.\-]?\d{2}){4}(?!\d)';  Confidence = "Basse";   Category = "Telephone" },
  @{ Name = "Tel international"; Pattern = '(?<!\d)\+(?:3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9])\s?\d(?:[\s.\-]?\d){7,12}(?!\d)'; Confidence = "Moyenne"; Category = "Telephone" }
)

$regexIBAN = @(
  @{ Name = "IBAN FR";            Pattern = '(?<!\w)FR\s?\d{2}[\s.\-]?(?:\d{4}[\s.\-]?){5}\d{3}(?!\w)'; Confidence = "Haute"; Category = "IBAN" },
  # IBAN international durci v2.3 :
  #   - (?!FR) evite les doublons avec la regle IBAN FR
  #   - \d{2} sans espace apres le code pays (un vrai IBAN colle les chiffres de controle)
  #   - {3,7} exige au minimum 4 blocs de 4 = 16 caracteres apres le code pays
  #     => elimine les faux positifs courts (VS3307, OK29, PP1055, PM1559, etc.)
  @{ Name = "IBAN international"; Pattern = '(?<!\w)(?!FR)[A-Z]{2}\d{2}[\s.\-]?[A-Z0-9]{4}(?:[\s.\-]?[A-Z0-9]{4}){3,7}(?:[\s.\-]?[A-Z0-9]{1,4})?(?!\w)'; Confidence = "Moyenne"; Category = "IBAN" }
)

$regexNIR = @(
  @{ Name = "NIR avec cle"; Pattern = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{2}(?!\d)'; Confidence = "Haute"; Category = "Securite Sociale" },
  @{ Name = "NIR sans cle"; Pattern = '(?<!\d)[12478]\s?\d{2}[\s.\-]?(?:0[1-9]|1[0-2]|[2-4]\d|5[0-9]|[6-9]\d)[\s.\-]?(?:\d{2}|2[AB])[\s.\-]?\d{3}[\s.\-]?\d{3}(?!\d)'; Confidence = "Moyenne"; Category = "Securite Sociale" }
)

$regexAdresse = @(
  @{ Name = "Adresse avec voie";   Pattern = '(?<!\w)\d{1,4}[\s,]+(?:rue|avenue|boulevard|impasse|allee|place|chemin|route|passage|cours|square|residence|lotissement|hameau|lieu[\s-]?dit|av\.|bd\.|bld\.|r\.|pl\.)\s+[A-Za-z\u00C0-\u017F\s\-'']{3,40}(?!\w)'; Confidence = "Haute"; Category = "Adresse" },
  @{ Name = "Code postal + ville"; Pattern = '(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5]|97[1-6]|98[4-9])\d{3}\s+[A-Z\u00C0-\u017F][A-Za-z\u00C0-\u017F\s\-'']{2,30}(?!\w)'; Confidence = "Moyenne"; Category = "Adresse" }
)

$allRules = $regexTelFR + $regexIBAN + $regexNIR + $regexAdresse

# ============================================================
# PATTERNS D'EXCLUSION GENERIQUES (anti faux-positifs)
# ============================================================

$exclusionPatterns = @(
  '^\d{4}-\d{2}-\d{2}',                    # Dates ISO
  '^\d{2}/\d{2}/\d{4}',                    # Dates FR
  '^[A-Z]{2,10}-\d+',                      # Cles Jira
  '^v?\d+\.\d+\.\d+',                      # Versions
  '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}',  # Adresses IP
  '^https?://'                              # URLs
)

# ============================================================
# EXCLUSIONS ADRESSES D'ENTREPRISE (anti faux-positifs)
# ============================================================

# Mots-cles d'entreprise : si presents dans le match ou le contexte,
# la trouvaille est consideree comme adresse professionnelle connue.
# La comparaison est insensible a la casse (OrdinalIgnoreCase).
$excludedAddressKeywords = @(
  "Harmonie Mutuelle",
  "Mutex",
  "VYV",
  "Groupe VYV",
  "MGEN",
  "MNT",
  "Chorum",
  "SIHM",
  "Chatillon",
  "Châtillon"
)

# Adresses physiques connues des sites de l'entreprise.
# Recherche partielle insensible a la casse dans le match ET le contexte.
$excludedAddressPatterns = @(
  # Siege / sites Harmonie Mutuelle & Mutex
  "rue Blomet",
  "place Charles de Gaulle",
  "cours des 50 Otages",
  "cours des 50-Otages",
  "rue de Chateaudun",
  "rue de Châteaudun",
  "boulevard de Pesaro",
  "bd de Pesaro",
  "avenue du Marechal Juin",
  "avenue du Maréchal Juin",
  "av. du Marechal Juin",
  "av. du Maréchal Juin",
  "rue Francois Jacob",
  "rue François Jacob",
  "place Robert Schuman",
  "rue du Faubourg Saint-Honore",
  "rue du Faubourg Saint-Honoré",
  # Site de Chatillon
  "avenue de la Republique",
  "avenue de la République",
  "av. de la Republique",
  "av. de la République",
  "140 avenue de la R",
  "140, avenue de la R",
  "140 av. de la R"
)

# ============================================================
# FONCTIONS DE DETECTION
# ============================================================

function Test-Exclusion([string]$matchValue, [string]$context = "") {
  # 1. Exclusions generiques (dates, cles Jira, IP, URLs, versions)
  foreach ($exPattern in $exclusionPatterns) {
    if ($matchValue -match $exPattern) { return $true }
  }

  # 2. Exclusions adresses d'entreprise : mots-cles dans le match ou le contexte
  $textToCheck = "$matchValue $context"
  foreach ($kw in $excludedAddressKeywords) {
    if ($textToCheck.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
      return $true
    }
  }

  # 3. Exclusions adresses d'entreprise : adresses physiques connues
  foreach ($addr in $excludedAddressPatterns) {
    if ($matchValue.IndexOf($addr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    if ($context.IndexOf($addr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)    { return $true }
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
      return ($cle -eq (97 - ($base % 97)))
    } catch { return $false }
  }
  return $true
}

function Get-Context([string]$text, [int]$index, [int]$length, [int]$contextSize = 50) {
  $start = [Math]::Max(0, $index - $contextSize)
  $end   = [Math]::Min($text.Length, $index + $length + $contextSize)
  $ctx   = $text.Substring($start, $end - $start) -replace '[\r\n]+', ' ' -replace '\s+', ' '
  if ($start -gt 0)         { $ctx = "..." + $ctx }
  if ($end -lt $text.Length) { $ctx = $ctx + "..." }
  return $ctx.Trim()
}

function Extract-TextFromADF($node) {
  if ($null -eq $node) { return "" }
  $text = ""
  if ($node.type -eq "text" -and $node.text) { $text += [string]$node.text + " " }
  if ($node.content) { foreach ($child in $node.content) { $text += Extract-TextFromADF $child } }
  return $text
}

function Scan-Text([string]$text, [string]$issueKey, [string]$fieldName) {
  if (-not $text -or $text.Length -lt 5) { return @() }
  $findings = @()

  foreach ($rule in $allRules) {
    $matches = [regex]::Matches($text, $rule.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $matches) {
      $matchValue = $m.Value.Trim()

      # Calculer le contexte UNE SEULE FOIS (reutilise pour exclusion + rapport)
      $ctx = Get-Context $text $m.Index $m.Length

      # Test d'exclusion avec le contexte
      if (Test-Exclusion $matchValue $ctx) { continue }

      $confidence = $rule.Confidence

      # Validation specifique NIR
      if ($rule.Name -like "NIR*") {
        if (-not (Test-NIRChecksum $matchValue)) { $confidence = "Basse" }
      }

      # Validation specifique IBAN FR
      if ($rule.Name -eq "IBAN FR") {
        $ibanClean = $matchValue -replace '[\s.\-]', ''
        if ($ibanClean.Length -ne 27) { $confidence = "Basse" }
      }

      # Validation specifique IBAN international (v2.3)
      if ($rule.Name -eq "IBAN international") {
        $ibanClean = $matchValue -replace '[\s.\-]', ''
        # Un IBAN valide fait entre 15 (Norvege NO) et 34 caracteres
        if ($ibanClean.Length -lt 15 -or $ibanClean.Length -gt 34) { continue }
        # IBAN courts (< 18 car.) : confiance degradee car risque de faux positif
        if ($ibanClean.Length -lt 18) { $confidence = "Basse" }
      }

      $findings += [PSCustomObject]@{
        IssueKey   = $issueKey
        Field      = $fieldName
        Category   = $rule.Category
        RuleName   = $rule.Name
        MatchValue = $matchValue
        Confidence = $confidence
        Context    = $ctx
      }
    }
  }
  return $findings
}

# ============================================================
# BANNIERE
# ============================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  SCAN RGPD JIRA — DETECTION DONNEES SENSIBLES (v$VERSION)" -ForegroundColor Cyan
Write-Host "  Partie 1 : Champs texte (sans pieces jointes)" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Categories detectees :" -ForegroundColor White
Write-Host "    - Telephones (FR mobile/fixe/services, international)" -ForegroundColor DarkGray
Write-Host "    - IBAN (FR, international)" -ForegroundColor DarkGray
Write-Host "    - Numeros de Securite Sociale / NIR" -ForegroundColor DarkGray
Write-Host "    - Adresses postales (voie, code postal + ville)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Exclusions entreprise : Harmonie Mutuelle, Mutex, VYV, MGEN, MNT, Chorum, Chatillon" -ForegroundColor DarkGray
Write-Host ""

# --- Credentials ---
$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"

# ============================================================
# ETAPE 0 : RECUPERATION DES CATEGORIES DE PROJETS
# ============================================================

Write-Host "  Chargement des categories de projets..." -ForegroundColor DarkGray

$allProjects = @()
$startAt = 0
$projPageSize = 50
$hasMoreProjects = $true

while ($hasMoreProjects) {
  $url = "{0}/rest/api/3/project/search?startAt={1}&maxResults={2}&expand=projectCategory" -f $site.BaseUrl, $startAt, $projPageSize
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

  if (-not $resp.ok) {
    Write-Host ("  ERREUR : impossible de lister les projets (status={0})" -f $resp.status) -ForegroundColor Red
    return
  }

  $json = $resp.content | ConvertFrom-Json
  $values = $json.values
  $count = ($values | Measure-Object).Count

  if ($count -eq 0) { $hasMoreProjects = $false; break }

  foreach ($p in $values) {
    $catName = "Sans categorie"
    $catId   = "none"
    if ($p.projectCategory) {
      $catName = [string]$p.projectCategory.name
      $catId   = [string]$p.projectCategory.id
    }
    $allProjects += [PSCustomObject]@{
      Key          = [string]$p.key
      Name         = [string]$p.name
      ProjectType  = [string]$p.projectTypeKey
      CategoryName = $catName
      CategoryId   = $catId
    }
  }

  if ($json.isLast -eq $true) { $hasMoreProjects = $false }
  else { $startAt += $count }
}

Write-Host ("  {0} projets trouves sur le site." -f $allProjects.Count) -ForegroundColor Green

# Grouper par categorie
$categoryGroups = $allProjects | Group-Object CategoryName | Sort-Object Name

if ($categoryGroups.Count -eq 0) {
  Write-Host "  Aucune categorie trouvee." -ForegroundColor Red
  return
}

# ============================================================
# ETAPE 1 : SELECTION DE LA CATEGORIE
# ============================================================

if ($CategoryFilter) {
  $selectedCategory = $categoryGroups | Where-Object { $_.Name -like "*$CategoryFilter*" }
  if (-not $selectedCategory -or ($selectedCategory | Measure-Object).Count -eq 0) {
    Write-Host ("  ERREUR : aucune categorie correspondant a '{0}'" -f $CategoryFilter) -ForegroundColor Red
    Write-Host "  Categories disponibles :" -ForegroundColor Yellow
    foreach ($cg in $categoryGroups) {
      Write-Host ("    - {0} ({1} projets)" -f $cg.Name, $cg.Count) -ForegroundColor DarkGray
    }
    return
  }
  if (($selectedCategory | Measure-Object).Count -gt 1) {
    Write-Host ("  ERREUR : '{0}' correspond a plusieurs categories :" -f $CategoryFilter) -ForegroundColor Red
    foreach ($sc in $selectedCategory) {
      Write-Host ("    - {0} ({1} projets)" -f $sc.Name, $sc.Count) -ForegroundColor DarkGray
    }
    return
  }
  $selectedCategory = $selectedCategory[0]
} else {
  Write-Host "  CATEGORIES DE PROJETS DISPONIBLES :" -ForegroundColor White
  Write-Host ""
  $idx = 0
  foreach ($cg in $categoryGroups) {
    $idx++
    $projectList = ($cg.Group | ForEach-Object { $_.Key }) -join ", "
    if ($projectList.Length -gt 80) { $projectList = $projectList.Substring(0, 77) + "..." }
    Write-Host ("    {0,3}) {1,-35} [{2,3} projets]  {3}" -f $idx, $cg.Name, $cg.Count, $projectList) -ForegroundColor Yellow
  }
  Write-Host ""

  $choice = 0
  while ($choice -lt 1 -or $choice -gt $categoryGroups.Count) {
    $rawInput = Read-Host ("  Votre choix (1-{0})" -f $categoryGroups.Count)
    try { $choice = [int]$rawInput } catch { $choice = 0 }
  }
  $selectedCategory = $categoryGroups[$choice - 1]
}

$selectedProjects = $selectedCategory.Group
$categoryLabel = $selectedCategory.Name -replace '[^\w\-]', '_'

Write-Host ""
Write-Host ("  => Categorie selectionnee : {0}" -f $selectedCategory.Name) -ForegroundColor Green
Write-Host ("     {0} projets a scanner :" -f $selectedProjects.Count) -ForegroundColor Green
foreach ($sp in ($selectedProjects | Sort-Object Key)) {
  Write-Host ("       - {0,-12} {1}" -f $sp.Key, $sp.Name) -ForegroundColor DarkGray
}
Write-Host ""

$confirm = Read-Host "  Lancer le scan ? (O/N)"
if ($confirm -notmatch '^[OoYy]') {
  Write-Host "  Scan annule." -ForegroundColor Yellow
  return
}

# ============================================================
# FICHIERS DE SORTIE
# ============================================================

$script:logFile = Join-Path $ExportsDir ("ScanRGPD_{0}_{1}.log" -f $categoryLabel, $ts)
$script:csvFile = Join-Path $ExportsDir ("ScanRGPD_{0}_{1}.csv" -f $categoryLabel, $ts)
"" | Set-Content -Path $script:logFile -Encoding UTF8

Log "================================================================"
Log "  SCAN RGPD JIRA v$VERSION"
Log ("  Categorie : {0} ({1} projets)" -f $selectedCategory.Name, $selectedProjects.Count)
Log ("  Date      : {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Log ("  Site      : {0}" -f $site.BaseUrl)
Log "================================================================"

# ============================================================
# PREPARATION CSV
# ============================================================

$csvColumns = @("Ticket", "URL", "Projet", "Champ_PJ", "Categorie", "Regle", "Donnee_detectee", "Confiance", "Contexte")
Write-CsvHeader (($csvColumns | ForEach-Object { Escape-CsvField $_ }) -join ";")

# ============================================================
# SCAN PRINCIPAL — BOUCLE PAR PROJET
# ============================================================

$globalFindings  = New-Object System.Collections.Generic.List[object]
$globalScanned   = 0
$globalWithHits  = 0
$globalPages     = 0
$globalStartTime = Get-Date
$projectStats    = @{}
$projectIndex    = 0

foreach ($proj in ($selectedProjects | Sort-Object Key)) {

  $projectKey  = $proj.Key
  $projectName = $proj.Name
  $projectIndex++

  Log ""
  Log ("=== PROJET {0}/{1} : {2} ({3}) ===" -f $projectIndex, $selectedProjects.Count, $projectKey, $projectName)
  Write-Host ""
  Write-Host ("  [{0}/{1}] Projet : {2} ({3})" -f $projectIndex, $selectedProjects.Count, $projectKey, $projectName) -ForegroundColor Green

  # Pre-check
  $searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl
  $preCheckBody = @{
    jql        = "project = $projectKey"
    maxResults = 1
    fields     = @("key")
  } | ConvertTo-Json -Depth 5 -Compress

  $preCheckResp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $preCheckBody

  if (-not $preCheckResp.ok) {
    Log ("  Projet {0} inaccessible (status={1}), passage au suivant" -f $projectKey, $preCheckResp.status) "ERROR"
    Write-Host ("    ERREUR : inaccessible (status={0}), ignore" -f $preCheckResp.status) -ForegroundColor Red
    $projectStats[$projectKey] = @{ Name = $projectName; Scanned = 0; WithHits = 0; Findings = 0; Pages = 0; Duration = [TimeSpan]::Zero; Skipped = $true }
    continue
  }

  $preCheckJson = $preCheckResp.content | ConvertFrom-Json
  $preCheckCount = ($preCheckJson.issues | Measure-Object).Count

  if ($preCheckCount -eq 0) {
    Log ("  Projet {0} : aucun ticket" -f $projectKey) "WARN"
    Write-Host "    Aucun ticket, ignore." -ForegroundColor DarkGray
    $projectStats[$projectKey] = @{ Name = $projectName; Scanned = 0; WithHits = 0; Findings = 0; Pages = 0; Duration = [TimeSpan]::Zero; Skipped = $true }
    continue
  }

  # Compteurs par projet
  $pScanned = 0; $pWithHits = 0; $pFindings = 0; $pPages = 0
  $pStartTime = Get-Date
  $nextPageToken = $null
  $hasMore = $true
  $seenKeys = New-Object System.Collections.Generic.HashSet[string]

  while ($hasMore) {
    $pPages++
    $globalPages++

    $bodyObj = @{
      jql        = "project = $projectKey ORDER BY key ASC"
      maxResults = $PageSize
      fields     = @("key", "summary", "description", "comment")
    }
    if ($nextPageToken) { $bodyObj["nextPageToken"] = $nextPageToken }

    $resp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers `
            -Body ($bodyObj | ConvertTo-Json -Depth 5 -Compress)

    if (-not $resp.ok) {
      Log ("  Erreur API page {0} : status={1}" -f $pPages, $resp.status) "ERROR"
      break
    }

    $json   = $resp.content | ConvertFrom-Json
    $issues = $json.issues
    $issueCount = ($issues | Measure-Object).Count

    if ($issueCount -eq 0) { $hasMore = $false; break }

    # Detection de boucle
    $firstKey = [string]$issues[0].key
    if ($seenKeys.Contains($firstKey)) {
      Log ("  BOUCLE : {0} deja vu, arret force" -f $firstKey) "WARN"
      break
    }

    foreach ($issue in $issues) {
      $issueKey = [string]$issue.key
      [void]$seenKeys.Add($issueKey)
      $pScanned++; $globalScanned++
      $issueFindings = @()

      # --- Summary ---
      if ($issue.fields.summary) {
        $issueFindings += Scan-Text ([string]$issue.fields.summary) $issueKey "Summary"
      }

      # --- Description (ADF ou texte brut) ---
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
          $author = if ($comment.author.displayName) { [string]$comment.author.displayName } else { "?" }
          $issueFindings += Scan-Text $commentText $issueKey ("Commentaire #{0} ({1})" -f $cIdx, $author)
        }
      }

      # --- Champs custom texte ---
      foreach ($prop in $issue.fields.PSObject.Properties) {
        if ($prop.Name -like "customfield_*" -and $prop.Value) {
          $cfText = ""
          if ($prop.Value -is [string]) { $cfText = [string]$prop.Value }
          elseif ($prop.Value.type -eq "doc") { $cfText = Extract-TextFromADF $prop.Value }
          if ($cfText.Length -gt 5) {
            $issueFindings += Scan-Text $cfText $issueKey $prop.Name
          }
        }
      }

      # Ecrire dans le CSV au fil de l'eau
      if ($issueFindings.Count -gt 0) {
        $pWithHits++; $globalWithHits++
        foreach ($f in $issueFindings) {
          $pFindings++
          $globalFindings.Add($f) | Out-Null
          $issueUrl = "{0}/browse/{1}" -f $site.BaseUrl, $f.IssueKey
          $line = @(
            (Escape-CsvField $f.IssueKey),
            (Escape-CsvField $issueUrl),
            (Escape-CsvField $projectKey),
            (Escape-CsvField $f.Field),
            (Escape-CsvField $f.Category),
            (Escape-CsvField $f.RuleName),
            (Escape-CsvField $f.MatchValue),
            (Escape-CsvField $f.Confidence),
            (Escape-CsvField $f.Context)
          ) -join ";"
          Write-CsvLine $line
        }
      }
    }

    # Progression
    $elapsed = (Get-Date) - $pStartTime
    $rate = if ($elapsed.TotalSeconds -gt 0) { [Math]::Round($pScanned / $elapsed.TotalSeconds, 0) } else { 0 }
    $lastKey = [string]$issues[$issueCount - 1].key
    Write-Progress -Activity ("Scan {0} [{1}/{2}]" -f $projectKey, $projectIndex, $selectedProjects.Count) `
      -Status ("{0} scannes [{1}], {2} trouvailles ({3} t/s)" -f $pScanned, $lastKey, $pFindings, $rate) `
      -PercentComplete (-1)

    Log ("  Page {0} : {1} tickets [{2}..{3}], cumul={4}, hits={5}" -f $pPages, $issueCount, $firstKey, $lastKey, $pScanned, $pFindings)

    # Pagination
    if ($json.isLast -eq $true) { $hasMore = $false }
    elseif ($json.nextPageToken) { $nextPageToken = [string]$json.nextPageToken }
    else { $hasMore = $false }

    Start-Sleep -Milliseconds $ThrottleMs
  }

  Write-Progress -Activity ("Scan {0}" -f $projectKey) -Completed

  $pElapsed = (Get-Date) - $pStartTime
  $projectStats[$projectKey] = @{
    Name     = $projectName
    Scanned  = $pScanned
    WithHits = $pWithHits
    Findings = $pFindings
    Pages    = $pPages
    Duration = $pElapsed
    Skipped  = $false
  }

  Log ("  --- {0} : {1} tickets, {2} avec donnees, {3} trouvailles, {4:N1} min ---" -f `
    $projectKey, $pScanned, $pWithHits, $pFindings, $pElapsed.TotalMinutes)
  Write-Host ("    => {0} tickets, {1} trouvailles ({2:N1}s)" -f $pScanned, $pFindings, $pElapsed.TotalSeconds) -ForegroundColor $(if ($pFindings -gt 0) { "Yellow" } else { "Green" })
}

# ============================================================
# RAPPORT CONSOLE
# ============================================================

$globalElapsed = (Get-Date) - $globalStartTime

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  RESULTATS DU SCAN RGPD" -ForegroundColor Cyan
Write-Host ("  Categorie : {0}" -f $selectedCategory.Name) -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if ($globalFindings.Count -eq 0) {
  Write-Host "  Aucune donnee sensible detectee." -ForegroundColor Green
} else {

  # --- Par projet ---
  Write-Host "  PAR PROJET :" -ForegroundColor White
  $scannedProjects = $projectStats.GetEnumerator() | Where-Object { -not $_.Value.Skipped } | Sort-Object { $_.Value.Findings } -Descending
  foreach ($entry in $scannedProjects) {
    $ps = $entry.Value
    if ($ps.Findings -gt 0) {
      Write-Host ("    {0,-12} : {1,5} tickets, {2,4} trouvailles" -f $entry.Key, $ps.Scanned, $ps.Findings) -ForegroundColor Yellow
    } else {
      Write-Host ("    {0,-12} : {1,5} tickets, aucune trouvaille" -f $entry.Key, $ps.Scanned) -ForegroundColor DarkGray
    }
  }
  $skippedProjects = $projectStats.GetEnumerator() | Where-Object { $_.Value.Skipped }
  if (($skippedProjects | Measure-Object).Count -gt 0) {
    Write-Host ""
    Write-Host ("    ({0} projet(s) ignore(s) : vides ou inaccessibles)" -f ($skippedProjects | Measure-Object).Count) -ForegroundColor DarkGray
  }
  Write-Host ""

  # --- Par categorie de donnees ---
  Write-Host "  PAR CATEGORIE DE DONNEES :" -ForegroundColor White
  $catGroups = $globalFindings | Group-Object Category
  foreach ($cat in ($catGroups | Sort-Object Name)) {
    $color = switch ($cat.Name) {
      "Telephone"        { "Yellow" }
      "IBAN"             { "Red" }
      "Securite Sociale" { "Red" }
      "Adresse"          { "DarkYellow" }
      default            { "White" }
    }
    Write-Host ("    {0,-20} : {1}" -f $cat.Name, $cat.Count) -ForegroundColor $color
  }
  Write-Host ""

  # --- Par confiance ---
  Write-Host "  PAR NIVEAU DE CONFIANCE :" -ForegroundColor White
  $confGroups = $globalFindings | Group-Object Confidence
  foreach ($cg in ($confGroups | Sort-Object Name)) {
    $color = switch ($cg.Name) {
      "Haute"   { "Red" }
      "Moyenne" { "Yellow" }
      "Basse"   { "DarkGray" }
      default   { "White" }
    }
    $icon = switch ($cg.Name) { "Haute" { "[!]" } "Moyenne" { "[~]" } "Basse" { "[ ]" } default { "   " } }
    Write-Host ("    {0} {1,-10} : {2}" -f $icon, $cg.Name, $cg.Count) -ForegroundColor $color
  }
  Write-Host ""

  # --- Top 10 tickets ---
  $topIssues = $globalFindings | Group-Object IssueKey | Sort-Object Count -Descending | Select-Object -First 10
  Write-Host "  TOP 10 TICKETS LES PLUS EXPOSES :" -ForegroundColor White
  $rank = 0
  foreach ($ti in $topIssues) {
    $rank++
    $cats = ($ti.Group | ForEach-Object { $_.Category } | Select-Object -Unique) -join ", "
    Write-Host ("    {0,2}. {1,-15} : {2} trouvailles ({3})" -f $rank, $ti.Name, $ti.Count, $cats) -ForegroundColor Yellow
  }
}

# ============================================================
# RESUME FINAL
# ============================================================

$scannedCount = ($projectStats.GetEnumerator() | Where-Object { -not $_.Value.Skipped } | Measure-Object).Count
$skippedCount = ($projectStats.GetEnumerator() | Where-Object { $_.Value.Skipped } | Measure-Object).Count

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ("  Categorie             : {0}" -f $selectedCategory.Name)
Write-Host ("  Projets scannes       : {0} (+ {1} ignores)" -f $scannedCount, $skippedCount)
Write-Host ("  Tickets scannes       : {0}" -f $globalScanned)
Write-Host ("  Pages API             : {0}" -f $globalPages)
Write-Host ("  Tickets avec donnees  : {0}" -f $globalWithHits) `
  -ForegroundColor $(if ($globalWithHits -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Trouvailles totales   : {0}" -f $globalFindings.Count) `
  -ForegroundColor $(if ($globalFindings.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Duree totale          : {0:N1} min" -f $globalElapsed.TotalMinutes)
Write-Host ""
Write-Host ("  CSV : {0}" -f $script:csvFile) -ForegroundColor Green
Write-Host ("  LOG : {0}" -f $script:logFile) -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Recommandations
$hauteCount = ($globalFindings | Where-Object { $_.Confidence -eq "Haute" } | Measure-Object).Count
if ($hauteCount -gt 0) {
  Write-Host "  RECOMMANDATIONS :" -ForegroundColor Red
  Write-Host ("    - {0} trouvaille(s) HAUTE confiance => action urgente d'anonymisation" -f $hauteCount) -ForegroundColor Red
  Write-Host "    - Verifier manuellement les trouvailles MOYENNE confiance" -ForegroundColor Yellow
  Write-Host "    - Sensibiliser les equipes aux bonnes pratiques RGPD" -ForegroundColor White
  Write-Host ""
}

Log "Termine."