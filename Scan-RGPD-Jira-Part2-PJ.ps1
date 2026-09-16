<#
.SYNOPSIS
  Scan-RGPD-Jira-Part2-PJ.ps1
  Detection en masse de donnees personnelles sensibles dans les pieces jointes Jira.
  Partie 2 : Scan des PJ (PDF, DOCX, XLSX, images OCR, texte brut).

.DESCRIPTION
  Ce script :
  1. Recupere les categories de projets Jira du site
  2. Propose un menu de selection par categorie
  3. Pour chaque ticket ayant des PJ, liste les pieces jointes
  4. Telecharge et extrait le texte selon le type de fichier
  5. Applique les memes regex et exclusions que la Partie 1
  6. Affiche chaque trouvaille en temps reel dans le log
  7. Genere un CSV (UTF-8 BOM, separateur ;) et un log

  PROTECTIONS PERFORMANCE v1.3 :
    - Taille max par PJ : 1 Mo par defaut
    - XLSX : plafond 5000 cellules, skip si sharedStrings > 2 Mo
    - Filtre par nom de fichier (liste_des_contrats, export_, dump_, etc.)
    - Timeout 30s sur API et telechargements
    - Logs de diagnostic avant/apres chaque appel
    - Affichage temps reel des trouvailles

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 1.3 - Partie 2 (PJ uniquement)
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Scan-RGPD-Jira-Part2-PJ.ps1
  .\Scan-RGPD-Jira-Part2-PJ.ps1 -CategoryFilter "software"
  .\Scan-RGPD-Jira-Part2-PJ.ps1 -CategoryFilter "software" -MaxFileSizeMB 2
  .\Scan-RGPD-Jira-Part2-PJ.ps1 -ApiTimeoutSec 60 -DownloadTimeoutSec 60
#>

[CmdletBinding()]
param(
  [string] $CategoryFilter     = "",
  [int]    $PageSize            = 100,
  [int]    $MaxRetries          = 5,
  [int]    $ThrottleMs          = 200,
  [int]    $MaxFileSizeMB       = 1,
  [int]    $ApiTimeoutSec       = 30,
  [int]    $DownloadTimeoutSec  = 30,
  [int]    $MaxXlsxCells        = 5000
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# ============================================================
# CONSTANTES
# ============================================================

$VERSION = "1.3-Part2-PJ"

$textExtensions  = @(".txt", ".csv", ".log", ".xml", ".json", ".html", ".htm", ".md", ".rtf")
$pdfExtensions   = @(".pdf")
$docxExtensions  = @(".docx")
$xlsxExtensions  = @(".xlsx")
$imageExtensions = @(".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".gif")

$allowedExtensions = $textExtensions + $pdfExtensions + $docxExtensions + $xlsxExtensions + $imageExtensions

$MaxFileSizeBytes = $MaxFileSizeMB * 1024 * 1024

# Patterns de noms de fichiers a exclure (listes, exports, dumps)
$skipFileNamePatterns = @(
  "liste_des_contrats",
  "liste_contrats",
  "export_contrats",
  "export_",
  "extraction_",
  "dump_",
  "backup_",
  "sauvegarde_",
  "listing_",
  "inventaire_",
  "base_adherents",
  "base_contrats",
  "fichier_national"
)

# ============================================================
# INITIALISATION
# ============================================================

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
$TempDir    = Join-Path $ExportsDir ("temp_pj_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }
if (-not (Test-Path $TempDir))    { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }

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
# HTTP HELPER (avec retry exponentiel + timeout + logs)
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
        TimeoutSec      = $ApiTimeoutSec
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
      $isTimeout = $_.Exception.Message -match "timed?\s*out|timeout" -or
                   ($_.Exception.InnerException -and $_.Exception.InnerException.Message -match "timed?\s*out|timeout")
      if ($isTimeout) {
        Log ("  TIMEOUT {0} (tentative {1}/{2})" -f $Method, $attempt, $MaxRetries) "WARN"
      }
      if ($attempt -gt $MaxRetries) {
        $reason = if ($isTimeout) { "TIMEOUT apres $MaxRetries tentatives" } else { $_.Exception.Message }
        return @{ ok = $false; status = $status; error = $reason; body = $errBody }
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

function Download-File([string]$Url, [string]$DestPath, [hashtable]$Headers) {
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      Invoke-WebRequest -Uri $Url -Headers $Headers -OutFile $DestPath `
        -UseBasicParsing -TimeoutSec $DownloadTimeoutSec -ErrorAction Stop
      return $true
    } catch {
      $isTimeout = $_.Exception.Message -match "timed?\s*out|timeout"
      if ($isTimeout) { Log ("    TIMEOUT download (tentative {0}/{1})" -f $attempt, $MaxRetries) "WARN" }
      if ($attempt -gt $MaxRetries) { return $false }
      $sleepSec = [Math]::Min(30, [Math]::Pow(2, [Math]::Min(4, $attempt)))
      Start-Sleep -Seconds $sleepSec
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
# DETECTION DES OUTILS EXTERNES
# ============================================================

function Find-ExternalTool([string]$ExeName, [string[]]$SearchPaths) {
  $inPath = Get-Command $ExeName -ErrorAction SilentlyContinue
  if ($inPath) { return $inPath.Source }
  foreach ($sp in $SearchPaths) {
    $candidate = Join-Path $sp $ExeName
    if (Test-Path $candidate) { return $candidate }
  }
  return $null
}

$tesseractPaths = @(
  "C:\Program Files\Tesseract-OCR",
  "C:\Program Files (x86)\Tesseract-OCR",
  "$env:LOCALAPPDATA\Tesseract-OCR",
  "$env:LOCALAPPDATA\Programs\Tesseract-OCR"
)

$popplerPaths = @(
  "C:\Program Files\poppler\Library\bin",
  "C:\Program Files (x86)\poppler\Library\bin",
  "C:\Program Files\poppler\bin",
  "$env:LOCALAPPDATA\poppler\Library\bin"
)
$popplerPathsResolved = @()
foreach ($pp in $popplerPaths) {
  if ($pp -match '\*') {
    $resolved = Get-Item $pp -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    if ($resolved) { $popplerPathsResolved += $resolved }
  } else {
    $popplerPathsResolved += $pp
  }
}
$dynamicPoppler = Get-Item "C:\Program Files\poppler-*\Library\bin" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
if ($dynamicPoppler) { $popplerPathsResolved += $dynamicPoppler }

$script:tesseractExe = Find-ExternalTool "tesseract.exe" $tesseractPaths
$script:pdftotextExe = Find-ExternalTool "pdftotext.exe" $popplerPathsResolved

$script:hasTesseract = $null -ne $script:tesseractExe
$script:hasPdftotext = $null -ne $script:pdftotextExe

# ============================================================
# REGEX DE DETECTION (identiques a la Partie 1 v2.3)
# ============================================================

$regexTelFR = @(
  @{ Name = "Tel FR mobile";     Pattern = '(?<!\d)(?:\+33\s?|0)(?:6|7)(?:[\s.\-]?\d{2}){4}(?!\d)';  Confidence = "Haute";   Category = "Telephone" },
  @{ Name = "Tel FR fixe";       Pattern = '(?<!\d)(?:\+33\s?|0)(?:1|2|3|4|5)(?:[\s.\-]?\d{2}){4}(?!\d)'; Confidence = "Moyenne"; Category = "Telephone" },
  @{ Name = "Tel FR services";   Pattern = '(?<!\d)(?:\+33\s?|0)(?:8|9)(?:[\s.\-]?\d{2}){4}(?!\d)';  Confidence = "Basse";   Category = "Telephone" },
  @{ Name = "Tel international"; Pattern = '(?<!\d)\+(?:3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9])\s?\d(?:[\s.\-]?\d){7,12}(?!\d)'; Confidence = "Moyenne"; Category = "Telephone" }
)

$regexIBAN = @(
  @{ Name = "IBAN FR";            Pattern = '(?<!\w)FR\s?\d{2}[\s.\-]?(?:\d{4}[\s.\-]?){5}\d{3}(?!\w)'; Confidence = "Haute"; Category = "IBAN" },
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
# PATTERNS D'EXCLUSION (identiques a la Partie 1 v2.3)
# ============================================================

$exclusionPatterns = @(
  '^\d{4}-\d{2}-\d{2}',
  '^\d{2}/\d{2}/\d{4}',
  '^[A-Z]{2,10}-\d+',
  '^v?\d+\.\d+\.\d+',
  '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}',
  '^https?://'
)

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

$excludedAddressPatterns = @(
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
  "avenue de la Republique",
  "avenue de la République",
  "av. de la Republique",
  "av. de la République",
  "140 avenue de la R",
  "140, avenue de la R",
  "140 av. de la R"
)

# ============================================================
# FONCTIONS DE DETECTION (identiques a la Partie 1 v2.3)
# ============================================================

function Test-Exclusion([string]$matchValue, [string]$context = "") {
  foreach ($exPattern in $exclusionPatterns) {
    if ($matchValue -match $exPattern) { return $true }
  }
  $textToCheck = "$matchValue $context"
  foreach ($kw in $excludedAddressKeywords) {
    if ($textToCheck.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
  }
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

function Scan-Text([string]$text, [string]$issueKey, [string]$fieldName) {
  if (-not $text -or $text.Length -lt 5) { return @() }
  $findings = @()
  foreach ($rule in $allRules) {
    $matches = [regex]::Matches($text, $rule.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $matches) {
      $matchValue = $m.Value.Trim()
      $ctx = Get-Context $text $m.Index $m.Length
      if (Test-Exclusion $matchValue $ctx) { continue }
      $confidence = $rule.Confidence
      if ($rule.Name -like "NIR*") {
        if (-not (Test-NIRChecksum $matchValue)) { $confidence = "Basse" }
      }
      if ($rule.Name -eq "IBAN FR") {
        $ibanClean = $matchValue -replace '[\s.\-]', ''
        if ($ibanClean.Length -ne 27) { $confidence = "Basse" }
      }
      if ($rule.Name -eq "IBAN international") {
        $ibanClean = $matchValue -replace '[\s.\-]', ''
        if ($ibanClean.Length -lt 15 -or $ibanClean.Length -gt 34) { continue }
        if ($ibanClean.Length -lt 18) { $confidence = "Basse" }
      }
      $findings += [PSCustomObject]@{
        IssueKey = $issueKey; Field = $fieldName; Category = $rule.Category
        RuleName = $rule.Name; MatchValue = $matchValue; Confidence = $confidence; Context = $ctx
      }
    }
  }
  return $findings
}

# ============================================================
# FONCTIONS D'EXTRACTION DE TEXTE
# ============================================================

function Extract-TextFromPlainFile([string]$filePath) {
  try {
    return [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::UTF8)
  } catch {
    try { return Get-Content -Path $filePath -Raw -Encoding Default } catch { return "" }
  }
}

function Extract-TextFromPdf([string]$filePath) {
  if (-not $script:hasPdftotext) { return $null }
  try {
    $outFile = $filePath + ".txt"
    $proc = Start-Process -FilePath $script:pdftotextExe `
      -ArgumentList "-layout", "-enc", "UTF-8", "`"$filePath`"", "`"$outFile`"" `
      -NoNewWindow -Wait -PassThru -ErrorAction Stop
    if ($proc.ExitCode -eq 0 -and (Test-Path $outFile)) {
      $text = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8)
      Remove-Item $outFile -Force -ErrorAction SilentlyContinue
      return $text
    }
    Remove-Item $outFile -Force -ErrorAction SilentlyContinue
    return ""
  } catch { return "" }
}

function Extract-TextFromDocx([string]$filePath) {
  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
    $text = ""
    foreach ($entry in $zip.Entries) {
      if ($entry.FullName -eq "word/document.xml") {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close(); $stream.Close()
        $xmlDoc = [xml]$xmlContent
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $nodes = $xmlDoc.SelectNodes("//w:t", $nsMgr)
        foreach ($node in $nodes) { $text += $node.InnerText + " " }
        break
      }
    }
    $zip.Dispose()
    return $text
  } catch { return "" }
}

function Extract-TextFromXlsx([string]$filePath) {
  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
    $text = New-Object System.Text.StringBuilder
    $cellCount = 0

    # 1. Charger les shared strings (avec garde-fou taille)
    $sharedStrings = @()
    foreach ($entry in $zip.Entries) {
      if ($entry.FullName -eq "xl/sharedStrings.xml") {
        # Skip si le fichier sharedStrings est trop gros
        if ($entry.Length -gt 2 * 1024 * 1024) {
          Log ("      XLSX skip : sharedStrings trop gros ({0:N0} Ko)" -f ($entry.Length / 1KB)) "WARN"
          $zip.Dispose()
          return ""
        }
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close(); $stream.Close()
        $xmlDoc = [xml]$xmlContent
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsMgr.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $siNodes = $xmlDoc.SelectNodes("//s:si", $nsMgr)
        foreach ($si in $siNodes) {
          $tNodes = $si.SelectNodes(".//s:t", $nsMgr)
          $cellText = ""
          foreach ($t in $tNodes) { $cellText += $t.InnerText }
          $sharedStrings += $cellText
        }
        break
      }
    }

    # 2. Lire les feuilles (avec plafond de cellules)
    foreach ($entry in $zip.Entries) {
      if ($cellCount -ge $MaxXlsxCells) { break }
      if ($entry.FullName -match "^xl/worksheets/sheet\d+\.xml$") {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close(); $stream.Close()
        $xmlDoc = [xml]$xmlContent
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsMgr.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $cells = $xmlDoc.SelectNodes("//s:c", $nsMgr)
        foreach ($cell in $cells) {
          if ($cellCount -ge $MaxXlsxCells) {
            Log ("      XLSX plafond atteint : {0} cellules max" -f $MaxXlsxCells) "WARN"
            break
          }
          $vNode = $cell.SelectSingleNode("s:v", $nsMgr)
          if ($vNode) {
            $cellType = $cell.GetAttribute("t")
            if ($cellType -eq "s" -and $sharedStrings.Count -gt 0) {
              $idx = [int]$vNode.InnerText
              if ($idx -lt $sharedStrings.Count) { [void]$text.Append($sharedStrings[$idx] + " ") }
            } else {
              [void]$text.Append($vNode.InnerText + " ")
            }
            $cellCount++
          }
        }
      }
    }
    $zip.Dispose()
    return $text.ToString()
  } catch { return "" }
}

function Extract-TextFromImage([string]$filePath) {
  if (-not $script:hasTesseract) { return $null }
  try {
    $outBase = $filePath + "_ocr"
    $outFile = $outBase + ".txt"
    $proc = Start-Process -FilePath $script:tesseractExe `
      -ArgumentList "`"$filePath`"", "`"$outBase`"", "-l", "fra+eng", "--psm", "3" `
      -NoNewWindow -Wait -PassThru -ErrorAction Stop
    if ($proc.ExitCode -eq 0 -and (Test-Path $outFile)) {
      $text = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8)
      Remove-Item $outFile -Force -ErrorAction SilentlyContinue
      return $text
    }
    Remove-Item $outFile -Force -ErrorAction SilentlyContinue
    return ""
  } catch { return "" }
}

function Extract-TextFromFile([string]$filePath, [string]$extension) {
  switch -Wildcard ($extension.ToLower()) {
    { $_ -in $textExtensions }  { return Extract-TextFromPlainFile $filePath }
    ".pdf"                       { return Extract-TextFromPdf $filePath }
    ".docx"                      { return Extract-TextFromDocx $filePath }
    ".xlsx"                      { return Extract-TextFromXlsx $filePath }
    { $_ -in $imageExtensions } { return Extract-TextFromImage $filePath }
    default                      { return $null }
  }
}

# ============================================================
# BANNIERE
# ============================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  SCAN RGPD JIRA — PIECES JOINTES (v$VERSION)" -ForegroundColor Cyan
Write-Host "  Partie 2 : Scan des PJ (PDF, DOCX, XLSX, images, texte)" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Extensions scannees :" -ForegroundColor White
Write-Host ("    Texte  : {0}" -f ($textExtensions -join ", ")) -ForegroundColor DarkGray
Write-Host ("    PDF    : {0}" -f ($pdfExtensions -join ", ")) -ForegroundColor DarkGray
Write-Host ("    Office : {0}" -f (($docxExtensions + $xlsxExtensions) -join ", ")) -ForegroundColor DarkGray
Write-Host ("    Images : {0}" -f ($imageExtensions -join ", ")) -ForegroundColor DarkGray
Write-Host ""
Write-Host ("  Taille max par PJ    : {0} Mo" -f $MaxFileSizeMB) -ForegroundColor DarkGray
Write-Host ("  Max cellules XLSX    : {0}" -f $MaxXlsxCells) -ForegroundColor DarkGray
Write-Host ("  Timeout API          : {0}s" -f $ApiTimeoutSec) -ForegroundColor DarkGray
Write-Host ("  Timeout download     : {0}s" -f $DownloadTimeoutSec) -ForegroundColor DarkGray
Write-Host ""

Write-Host "  Outils externes :" -ForegroundColor White
if ($script:hasPdftotext) {
  Write-Host ("    [OK] pdftotext : {0}" -f $script:pdftotextExe) -ForegroundColor Green
} else {
  Write-Host "    [!!] pdftotext : NON TROUVE — les PDF seront ignores" -ForegroundColor Yellow
  Write-Host "         Installer Poppler : winget install poppler" -ForegroundColor DarkGray
}
if ($script:hasTesseract) {
  Write-Host ("    [OK] Tesseract : {0}" -f $script:tesseractExe) -ForegroundColor Green
} else {
  Write-Host "    [!!] Tesseract : NON TROUVE — les images seront ignorees" -ForegroundColor Yellow
  Write-Host "         Installer : winget install UB-Mannheim.TesseractOCR" -ForegroundColor DarkGray
}
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
    $catName = "Sans categorie"; $catId = "none"
    if ($p.projectCategory) {
      $catName = [string]$p.projectCategory.name
      $catId   = [string]$p.projectCategory.id
    }
    $allProjects += [PSCustomObject]@{
      Key = [string]$p.key; Name = [string]$p.name
      ProjectType = [string]$p.projectTypeKey
      CategoryName = $catName; CategoryId = $catId
    }
  }
  if ($json.isLast -eq $true) { $hasMoreProjects = $false } else { $startAt += $count }
}

Write-Host ("  {0} projets trouves." -f $allProjects.Count) -ForegroundColor Green
$categoryGroups = $allProjects | Group-Object CategoryName | Sort-Object Name
if ($categoryGroups.Count -eq 0) { Write-Host "  Aucune categorie." -ForegroundColor Red; return }

# ============================================================
# ETAPE 1 : SELECTION DE LA CATEGORIE
# ============================================================

if ($CategoryFilter) {
  $selectedCategory = $categoryGroups | Where-Object { $_.Name -like "*$CategoryFilter*" }
  if (-not $selectedCategory -or ($selectedCategory | Measure-Object).Count -eq 0) {
    Write-Host ("  ERREUR : aucune categorie correspondant a '{0}'" -f $CategoryFilter) -ForegroundColor Red
    foreach ($cg in $categoryGroups) { Write-Host ("    - {0} ({1} projets)" -f $cg.Name, $cg.Count) -ForegroundColor DarkGray }
    return
  }
  if (($selectedCategory | Measure-Object).Count -gt 1) {
    Write-Host ("  ERREUR : '{0}' correspond a plusieurs categories" -f $CategoryFilter) -ForegroundColor Red
    foreach ($sc in $selectedCategory) { Write-Host ("    - {0} ({1} projets)" -f $sc.Name, $sc.Count) -ForegroundColor DarkGray }
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
Write-Host ("  => Categorie : {0} ({1} projets)" -f $selectedCategory.Name, $selectedProjects.Count) -ForegroundColor Green
foreach ($sp in ($selectedProjects | Sort-Object Key)) {
  Write-Host ("       - {0,-12} {1}" -f $sp.Key, $sp.Name) -ForegroundColor DarkGray
}
Write-Host ""

$confirm = Read-Host "  Lancer le scan des PJ ? (O/N)"
if ($confirm -notmatch '^[OoYy]') { Write-Host "  Scan annule." -ForegroundColor Yellow; return }

# ============================================================
# FICHIERS DE SORTIE
# ============================================================

$script:logFile = Join-Path $ExportsDir ("ScanRGPD-PJ_{0}_{1}.log" -f $categoryLabel, $ts)
$script:csvFile = Join-Path $ExportsDir ("ScanRGPD-PJ_{0}_{1}.csv" -f $categoryLabel, $ts)
"" | Set-Content -Path $script:logFile -Encoding UTF8

Log "================================================================"
Log "  SCAN RGPD JIRA — PIECES JOINTES v$VERSION"
Log ("  Categorie : {0} ({1} projets)" -f $selectedCategory.Name, $selectedProjects.Count)
Log ("  Date      : {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Log ("  Site      : {0}" -f $site.BaseUrl)
Log ("  pdftotext : {0}" -f $(if ($script:hasPdftotext) { $script:pdftotextExe } else { "NON DISPONIBLE" }))
Log ("  Tesseract : {0}" -f $(if ($script:hasTesseract) { $script:tesseractExe } else { "NON DISPONIBLE" }))
Log ("  Params    : MaxPJ={0}Mo, MaxCells={1}, TimeoutAPI={2}s, TimeoutDL={3}s" -f $MaxFileSizeMB, $MaxXlsxCells, $ApiTimeoutSec, $DownloadTimeoutSec)
Log "================================================================"

$csvColumns = @("Ticket", "URL", "Projet", "Champ_PJ", "Categorie", "Regle", "Donnee_detectee", "Confiance", "Contexte")
Write-CsvHeader (($csvColumns | ForEach-Object { Escape-CsvField $_ }) -join ";")

# ============================================================
# SCAN PRINCIPAL — BOUCLE PAR PROJET
# ============================================================

$globalFindings    = New-Object System.Collections.Generic.List[object]
$globalScanned     = 0
$globalWithHits    = 0
$globalPages       = 0
$globalPJScanned   = 0
$globalPJSkipped   = 0
$globalPJNoTool    = 0
$globalPJErrors    = 0
$globalStartTime   = Get-Date
$projectStats      = @{}
$projectIndex      = 0

foreach ($proj in ($selectedProjects | Sort-Object Key)) {

  $projectKey  = $proj.Key
  $projectName = $proj.Name
  $projectIndex++

  Log ""
  Log ("=== PROJET {0}/{1} : {2} ({3}) ===" -f $projectIndex, $selectedProjects.Count, $projectKey, $projectName)
  Write-Host ""
  Write-Host ("  [{0}/{1}] Projet : {2} ({3})" -f $projectIndex, $selectedProjects.Count, $projectKey, $projectName) -ForegroundColor Green

  $searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl

  # --- Pre-check avec log ---
  Log ("  >> Pre-check projet {0}..." -f $projectKey)
  Write-Host ("    Pre-check {0}..." -f $projectKey) -ForegroundColor DarkGray -NoNewline

  $preCheckBody = @{
    jql = "project = $projectKey AND attachments IS NOT EMPTY"
    maxResults = 1
    fields = @("key")
  } | ConvertTo-Json -Depth 5 -Compress

  $preCheckResp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $preCheckBody
  Log ("  << Pre-check termine (status={0})" -f $(if ($preCheckResp.ok) { "OK" } else { $preCheckResp.status }))

  if (-not $preCheckResp.ok) {
    Write-Host " ERREUR" -ForegroundColor Red
    if ($preCheckResp.status -eq 400) {
      Log ("  ERREUR 400 sur projet {0} : JQL ou schema incompatible" -f $projectKey) "ERROR"
      Write-Host ("    ERREUR 400 : projet {0} incompatible, ignore" -f $projectKey) -ForegroundColor Red
    } elseif ($preCheckResp.error -match "TIMEOUT") {
      Log ("  TIMEOUT sur pre-check projet {0}" -f $projectKey) "ERROR"
      Write-Host ("    TIMEOUT : projet {0} ne repond pas, ignore" -f $projectKey) -ForegroundColor Red
    } else {
      Log ("  Projet {0} inaccessible (status={1})" -f $projectKey, $preCheckResp.status) "ERROR"
      Write-Host ("    ERREUR status={0}, ignore" -f $preCheckResp.status) -ForegroundColor Red
    }
    $projectStats[$projectKey] = @{
      Name = $projectName; Scanned = 0; WithHits = 0; Findings = 0
      PJScanned = 0; PJSkipped = 0; PJNoTool = 0; PJErrors = 0
      Pages = 0; Duration = [TimeSpan]::Zero; Skipped = $true
    }
    continue
  }

  Write-Host " OK" -ForegroundColor Green

  $preCheckJson = $preCheckResp.content | ConvertFrom-Json
  if (($preCheckJson.issues | Measure-Object).Count -eq 0) {
    Log ("  Projet {0} : aucun ticket avec PJ" -f $projectKey) "WARN"
    Write-Host "    Aucun ticket avec PJ, ignore." -ForegroundColor DarkGray
    $projectStats[$projectKey] = @{
      Name = $projectName; Scanned = 0; WithHits = 0; Findings = 0
      PJScanned = 0; PJSkipped = 0; PJNoTool = 0; PJErrors = 0
      Pages = 0; Duration = [TimeSpan]::Zero; Skipped = $true
    }
    continue
  }

  $pScanned = 0; $pWithHits = 0; $pFindings = 0; $pPages = 0
  $pPJScanned = 0; $pPJSkipped = 0; $pPJNoTool = 0; $pPJErrors = 0
  $pStartTime = Get-Date
  $nextPageToken = $null
  $hasMore = $true
  $seenKeys = New-Object System.Collections.Generic.HashSet[string]

  while ($hasMore) {
    $pPages++; $globalPages++

    $bodyObj = @{
      jql        = "project = $projectKey AND attachments IS NOT EMPTY ORDER BY key ASC"
      maxResults = $PageSize
      fields     = @("key", "attachment")
    }
    if ($nextPageToken) { $bodyObj["nextPageToken"] = $nextPageToken }

    Log ("  >> Appel API page {0} (projet {1})..." -f $pPages, $projectKey)
    Write-Host ("    Requete page {0}..." -f $pPages) -ForegroundColor DarkGray -NoNewline

    $resp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers `
            -Body ($bodyObj | ConvertTo-Json -Depth 5 -Compress)

    if (-not $resp.ok) {
      Write-Host " ERREUR" -ForegroundColor Red
      if ($resp.status -eq 400) {
        Log ("  ERREUR 400 sur projet {0} page {1}, passage au projet suivant" -f $projectKey, $pPages) "ERROR"
        Write-Host ("    ERREUR 400 : projet {0} incompatible, ignore" -f $projectKey) -ForegroundColor Red
      } elseif ($resp.error -match "TIMEOUT") {
        Log ("  TIMEOUT sur projet {0} page {1}, passage au projet suivant" -f $projectKey, $pPages) "ERROR"
        Write-Host ("    TIMEOUT : projet {0} ne repond pas, ignore" -f $projectKey) -ForegroundColor Red
      } else {
        Log ("  Erreur API page {0} : status={1} — {2}" -f $pPages, $resp.status, $resp.error) "ERROR"
      }
      break
    }

    Write-Host " OK" -ForegroundColor Green

    $json = $resp.content | ConvertFrom-Json
    $issues = $json.issues
    $issueCount = ($issues | Measure-Object).Count
    if ($issueCount -eq 0) { $hasMore = $false; break }

    $firstKey = [string]$issues[0].key
    if ($seenKeys.Contains($firstKey)) { Log "  BOUCLE detectee, arret" "WARN"; break }

    foreach ($issue in $issues) {
      $issueKey = [string]$issue.key
      [void]$seenKeys.Add($issueKey)
      $pScanned++; $globalScanned++
      $issueHasHit = $false

      $attachments = $issue.fields.attachment
      if (-not $attachments) { continue }

      foreach ($att in $attachments) {
        $fileName = [string]$att.filename
        $fileSize = [long]$att.size
        $fileUrl  = [string]$att.content
        $ext      = [System.IO.Path]::GetExtension($fileName).ToLower()

        # Filtre extension
        if ($ext -notin $allowedExtensions) { continue }

        # Filtre par nom de fichier (listes, exports, dumps)
        $fileNameLower = $fileName.ToLower()
        $skipFile = $false
        foreach ($sp in $skipFileNamePatterns) {
          if ($fileNameLower.Contains($sp)) { $skipFile = $true; break }
        }
        if ($skipFile) {
          Log ("    SKIP {0} (nom exclu : liste/export/dump)" -f $fileName) "WARN"
          $pPJSkipped++; $globalPJSkipped++
          continue
        }

        # Filtre taille
        if ($fileSize -gt $MaxFileSizeBytes) {
          Log ("    SKIP {0} ({1:N0} Ko > {2} Mo)" -f $fileName, ($fileSize / 1KB), $MaxFileSizeMB) "WARN"
          $pPJSkipped++; $globalPJSkipped++
          continue
        }

        # Verifier outil disponible
        if ($ext -in $pdfExtensions -and -not $script:hasPdftotext) {
          $pPJNoTool++; $globalPJNoTool++; continue
        }
        if ($ext -in $imageExtensions -and -not $script:hasTesseract) {
          $pPJNoTool++; $globalPJNoTool++; continue
        }

        # --- Telecharger avec log ---
        $safeName = $fileName -replace '[^\w\.\-]', '_'
        $localPath = Join-Path $TempDir ("{0}_{1}" -f $issueKey, $safeName)

        Log ("    >> Download {0} ({1:N0} Ko)..." -f $fileName, ($fileSize / 1KB))
        $downloaded = Download-File -Url $fileUrl -DestPath $localPath -Headers $site.Headers

        if (-not $downloaded) {
          Log ("    ERREUR telechargement {0} ({1})" -f $fileName, $issueKey) "ERROR"
          $pPJErrors++; $globalPJErrors++
          continue
        }

        Log ("    << Download OK : {0}" -f $fileName)
        $pPJScanned++; $globalPJScanned++

        # --- Extraire le texte (avec mesure du temps) ---
        $extractedText = $null
        $extractStart = Get-Date
        try {
          $extractedText = Extract-TextFromFile $localPath $ext
        } catch {
          Log ("    ERREUR extraction {0} : {1}" -f $fileName, $_.Exception.Message) "ERROR"
          $pPJErrors++; $globalPJErrors++
        }
        $extractDuration = ((Get-Date) - $extractStart).TotalSeconds
        if ($extractDuration -gt 5) {
          Log ("    LENT : extraction {0} a pris {1:N1}s" -f $fileName, $extractDuration) "WARN"
        }

        # Nettoyer
        Remove-Item $localPath -Force -ErrorAction SilentlyContinue

        if (-not $extractedText -or $extractedText.Length -lt 5) { continue }

        # --- Scanner et afficher les trouvailles en temps reel ---
        $pjLabel = "PJ: {0}" -f $fileName
        $pjFindings = Scan-Text $extractedText $issueKey $pjLabel

        if ($pjFindings.Count -gt 0) {
          if (-not $issueHasHit) { $pWithHits++; $globalWithHits++; $issueHasHit = $true }
          foreach ($f in $pjFindings) {
            $pFindings++
            $globalFindings.Add($f) | Out-Null

            # === AFFICHAGE TEMPS REEL ===
            $hitColor = switch ($f.Confidence) { "Haute" { "Red" } "Moyenne" { "Yellow" } "Basse" { "DarkGray" } default { "White" } }
            $hitIcon  = switch ($f.Confidence) { "Haute" { "[!]" } "Moyenne" { "[~]" } "Basse" { "[ ]" } default { "   " } }
            Write-Host ("      {0} {1} | {2} | {3} | {4}" -f $hitIcon, $f.IssueKey, $f.Category, $f.RuleName, $f.MatchValue) -ForegroundColor $hitColor
            Log ("    TROUVAILLE : {0} | {1} | {2} | {3} | {4} | Confiance={5}" -f $f.IssueKey, $f.Field, $f.Category, $f.RuleName, $f.MatchValue, $f.Confidence) "FOUND"

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
    }

    # Progression
    $elapsed = (Get-Date) - $pStartTime
    $rate = if ($elapsed.TotalSeconds -gt 0) { [Math]::Round($pScanned / $elapsed.TotalSeconds, 0) } else { 0 }
    $lastKey = [string]$issues[$issueCount - 1].key
    Write-Progress -Activity ("Scan PJ {0} [{1}/{2}]" -f $projectKey, $projectIndex, $selectedProjects.Count) `
      -Status ("{0} tickets, {1} PJ, {2} trouvailles ({3} t/s)" -f $pScanned, $pPJScanned, $pFindings, $rate) `
      -PercentComplete (-1)

    Log ("  Page {0} : {1} tickets [{2}..{3}], PJ={4}, hits={5}" -f $pPages, $issueCount, $firstKey, $lastKey, $pPJScanned, $pFindings)

    # Pagination
    if ($json.isLast -eq $true) { $hasMore = $false }
    elseif ($json.nextPageToken) { $nextPageToken = [string]$json.nextPageToken }
    else { $hasMore = $false }

    Start-Sleep -Milliseconds $ThrottleMs
  }

  Write-Progress -Activity ("Scan PJ {0}" -f $projectKey) -Completed

  $pElapsed = (Get-Date) - $pStartTime
  $projectStats[$projectKey] = @{
    Name = $projectName; Scanned = $pScanned; WithHits = $pWithHits; Findings = $pFindings
    PJScanned = $pPJScanned; PJSkipped = $pPJSkipped; PJNoTool = $pPJNoTool; PJErrors = $pPJErrors
    Pages = $pPages; Duration = $pElapsed; Skipped = $false
  }

  Log ("  --- {0} : {1} tickets, {2} PJ, {3} trouvailles, {4:N1} min ---" -f `
    $projectKey, $pScanned, $pPJScanned, $pFindings, $pElapsed.TotalMinutes)
  Write-Host ("    => {0} tickets, {1} PJ, {2} trouvailles ({3:N1}s)" -f $pScanned, $pPJScanned, $pFindings, $pElapsed.TotalSeconds) `
    -ForegroundColor $(if ($pFindings -gt 0) { "Yellow" } else { "Green" })
}

# ============================================================
# NETTOYAGE TEMP
# ============================================================

try {
  $remainingFiles = Get-ChildItem $TempDir -ErrorAction SilentlyContinue
  if (($remainingFiles | Measure-Object).Count -eq 0) {
    Remove-Item $TempDir -Force -ErrorAction SilentlyContinue
    Log "  Repertoire temporaire supprime."
  } else {
    Log ("  {0} fichier(s) restant(s) dans {1}" -f ($remainingFiles | Measure-Object).Count, $TempDir) "WARN"
  }
} catch {}

# ============================================================
# RAPPORT CONSOLE
# ============================================================

$globalElapsed = (Get-Date) - $globalStartTime

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  RESULTATS DU SCAN RGPD — PIECES JOINTES" -ForegroundColor Cyan
Write-Host ("  Categorie : {0}" -f $selectedCategory.Name) -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if ($globalFindings.Count -eq 0) {
  Write-Host "  Aucune donnee sensible detectee dans les PJ." -ForegroundColor Green
} else {

  Write-Host "  PAR PROJET :" -ForegroundColor White
  $scannedProjects = $projectStats.GetEnumerator() | Where-Object { -not $_.Value.Skipped } | Sort-Object { $_.Value.Findings } -Descending
  foreach ($entry in $scannedProjects) {
    $ps = $entry.Value
    if ($ps.Findings -gt 0) {
      Write-Host ("    {0,-12} : {1,4} PJ scannees, {2,4} trouvailles" -f $entry.Key, $ps.PJScanned, $ps.Findings) -ForegroundColor Yellow
    } else {
      Write-Host ("    {0,-12} : {1,4} PJ scannees, aucune trouvaille" -f $entry.Key, $ps.PJScanned) -ForegroundColor DarkGray
    }
  }
  $skippedProjects = $projectStats.GetEnumerator() | Where-Object { $_.Value.Skipped }
  if (($skippedProjects | Measure-Object).Count -gt 0) {
    Write-Host ""
    Write-Host ("    ({0} projet(s) ignore(s) : vides, inaccessibles, erreur 400 ou timeout)" -f ($skippedProjects | Measure-Object).Count) -ForegroundColor DarkGray
  }
  Write-Host ""

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

  Write-Host "  PAR NIVEAU DE CONFIANCE :" -ForegroundColor White
  $confGroups = $globalFindings | Group-Object Confidence
  foreach ($cg in ($confGroups | Sort-Object Name)) {
    $color = switch ($cg.Name) {
      "Haute"   { "Red" }
      "Moyenne" { "Yellow" }
      "Basse"   { "DarkGray" }
      default   { "White" }
    }
    $icon = switch ($cg.Name) {
      "Haute"   { "[!]" }
      "Moyenne" { "[~]" }
      "Basse"   { "[ ]" }
      default   { "   " }
    }
    Write-Host ("    {0} {1,-10} : {2}" -f $icon, $cg.Name, $cg.Count) -ForegroundColor $color
  }
  Write-Host ""

  $topIssues = $globalFindings | Group-Object IssueKey | Sort-Object Count -Descending | Select-Object -First 10
  Write-Host "  TOP 10 TICKETS LES PLUS EXPOSES (PJ) :" -ForegroundColor White
  $rank = 0
  foreach ($ti in $topIssues) {
    $rank++
    $cats = ($ti.Group | ForEach-Object { $_.Category } | Select-Object -Unique) -join ", "
    $files = ($ti.Group | ForEach-Object { $_.Field } | Select-Object -Unique) -join ", "
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
Write-Host "  RESUME — PIECES JOINTES" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ("  Categorie             : {0}" -f $selectedCategory.Name)
Write-Host ("  Projets scannes       : {0} (+ {1} ignores)" -f $scannedCount, $skippedCount)
Write-Host ("  Tickets avec PJ       : {0}" -f $globalScanned)
Write-Host ("  PJ scannees           : {0}" -f $globalPJScanned) -ForegroundColor Green
Write-Host ("  PJ ignorees (taille)  : {0}" -f $globalPJSkipped) `
  -ForegroundColor $(if ($globalPJSkipped -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  PJ ignorees (outil)   : {0}" -f $globalPJNoTool) `
  -ForegroundColor $(if ($globalPJNoTool -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  PJ en erreur          : {0}" -f $globalPJErrors) `
  -ForegroundColor $(if ($globalPJErrors -gt 0) { "Red" } else { "Green" })
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

# Avertissement outils manquants
if ($globalPJNoTool -gt 0) {
  Write-Host "  OUTILS MANQUANTS :" -ForegroundColor Yellow
  if (-not $script:hasPdftotext) {
    Write-Host "    - Installer Poppler pour scanner les PDF : winget install poppler" -ForegroundColor Yellow
  }
  if (-not $script:hasTesseract) {
    Write-Host "    - Installer Tesseract pour l'OCR images : winget install UB-Mannheim.TesseractOCR" -ForegroundColor Yellow
  }
  Write-Host ""
}

# Recommandations
$hauteCount = ($globalFindings | Where-Object { $_.Confidence -eq "Haute" } | Measure-Object).Count
if ($hauteCount -gt 0) {
  Write-Host "  RECOMMANDATIONS :" -ForegroundColor Red
  Write-Host ("    - {0} trouvaille(s) HAUTE confiance dans les PJ => action urgente" -f $hauteCount) -ForegroundColor Red
  Write-Host "    - Verifier manuellement les PJ concernees" -ForegroundColor Yellow
  Write-Host "    - Envisager la suppression ou l'anonymisation des PJ" -ForegroundColor White
  Write-Host ""
}

Log "Termine."