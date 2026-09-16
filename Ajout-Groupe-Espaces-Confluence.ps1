<#
.SYNOPSIS
  Ajout-Groupe-Espaces-Confluence.ps1
  Ajoute un groupe DSIM avec un niveau de permission choisi
  sur une liste d'espaces Confluence (Jiradot).
.DESCRIPTION
  Ce script est entierement interactif :
  1. Ouvre un dialogue Windows pour selectionner le fichier de space keys
  2. Demande le mode (Dry / Execute)
  3. Liste les groupes Jira contenant "DSIM" sur Jiradot
  4. Propose un choix interactif du groupe
  5. Propose un choix de permission (Lecture / Ecriture / Admin / Ecriture+Suppression)
  6. Affiche le recapitulatif et applique (ou simule)

  CREDENTIALS :
    secrets\site-admin.xml => Jiradot (SiteUrl + Email + API Token)

  FICHIER D'ENTREE :
    Un fichier texte avec une space key par ligne (ou CSV avec colonne SpaceKey)
    Selectionne via un dialogue Windows

  FICHIERS GENERES (dans exports\) :
    AjoutGroupeConfluence_{ts}.csv   => Rapport des actions
    AjoutGroupeConfluence_{ts}.log   => Journal d execution
.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 2.0 - Correction API v1 POST permission
  Compatibilite  : PowerShell 5.1+
.EXAMPLE
  .\Ajout-Groupe-Espaces-Confluence.ps1
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
$logFile = Join-Path $ExportsDir ("AjoutGroupeConfluence_{0}.log" -f $ts)
$csvFile = Join-Path $ExportsDir ("AjoutGroupeConfluence_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
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

    $inputUrl = Read-Host "URL du site (ex: monsite.atlassian.net)"
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
# MAPPING DES PERMISSIONS CONFLUENCE
# ============================================================
$permissionSets = @{
  "Lecture" = @(
    @{ operation="read"; targetType="space" }
  )
  "Ecriture" = @(
    @{ operation="read"; targetType="space" },
    @{ operation="create"; targetType="page" },
    @{ operation="create"; targetType="blogpost" },
    @{ operation="create"; targetType="attachment" },
    @{ operation="create"; targetType="comment" }
  )
  "Ecriture+Suppression" = @(
    @{ operation="read"; targetType="space" },
    @{ operation="create"; targetType="page" },
    @{ operation="create"; targetType="blogpost" },
    @{ operation="create"; targetType="attachment" },
    @{ operation="create"; targetType="comment" },
    @{ operation="delete"; targetType="page" },
    @{ operation="delete"; targetType="blogpost" },
    @{ operation="delete"; targetType="attachment" },
    @{ operation="delete"; targetType="comment" }
  )
  "Admin" = @(
    @{ operation="read"; targetType="space" },
    @{ operation="administer"; targetType="space" }
  )
}

# ============================================================
# ETAPE 0 : DIALOGUE SELECTION DU FICHIER
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  AJOUT GROUPE DSIM SUR ESPACES CONFLUENCE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Selection du fichier de space keys..." -ForegroundColor White

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Selectionner le fichier de space keys Confluence"
$dialog.Filter = "Fichiers texte (*.txt;*.csv)|*.txt;*.csv|Tous les fichiers (*.*)|*.*"
$dialog.InitialDirectory = $ScriptDir
$dialog.Multiselect = $false
$result = $dialog.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Host "  Annule par l'utilisateur." -ForegroundColor Yellow
  return
}
$SpaceKeysFile = $dialog.FileName
Write-Host ("  => Fichier : {0}" -f $SpaceKeysFile) -ForegroundColor Green

# ============================================================
# ETAPE 0b : CHOIX DU MODE
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  MODE D'EXECUTION" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Dry     (simulation, aucune modification)" -ForegroundColor Yellow
Write-Host "  2. Execute (application reelle des permissions)" -ForegroundColor Yellow
Write-Host ""

$choixMode = 0
while ($choixMode -lt 1 -or $choixMode -gt 2) {
  $inputMode = Read-Host "  Choisissez le mode (1-2)"
  if ($inputMode -match "^\d+$") { $choixMode = [int]$inputMode }
}
$Mode = if ($choixMode -eq 1) { "Dry" } else { "Execute" }
Write-Host ""
Write-Host ("  => Mode : {0}" -f $Mode) -ForegroundColor Green

# ============================================================
# DEBUT DU TRAITEMENT
# ============================================================
Log "================================================================"
Log "  AJOUT GROUPE DSIM SUR ESPACES CONFLUENCE"
Log ("  Mode: {0}" -f $Mode)
Log ("  Fichier : {0}" -f $SpaceKeysFile)
Log "================================================================"

# --- Credentials ---
$site = Load-SiteCredentials -CredFile (Join-Path $SecretsDir "site-admin.xml") -SiteName "Jiradot"
Log ("  Site : {0}" -f $site.BaseUrl)

# ============================================================
# ETAPE 1 : LISTER LES GROUPES CONTENANT "DSIM"
# ============================================================
Log "=== ETAPE 1 : Recherche des groupes contenant DSIM ==="
$dsimGroups = New-Object System.Collections.Generic.List[object]
$startAt = 0; $pageSize = 100

while ($true) {
  $url = "{0}/rest/api/3/group/bulk?maxResults={1}&startAt={2}" -f $site.BaseUrl, $pageSize, $startAt
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers
  if (-not $resp.ok) {
    Log ("  Erreur listing groupes : status={0}" -f $resp.status) "ERROR"
    break
  }
  $json = $resp.content | ConvertFrom-Json
  foreach ($g in $json.values) {
    $gName = [string]$g.name
    if ($gName -match "DSIM") {
      $dsimGroups.Add(@{
        name    = $gName
        groupId = [string]$g.groupId
      }) | Out-Null
    }
  }
  if ($json.isLast -eq $true -or $json.values.Count -lt $pageSize) { break }
  $startAt += $pageSize
  Start-Sleep -Milliseconds 200
}

$cGroups = ($dsimGroups | Measure-Object).Count
Log ("  {0} groupes DSIM trouves" -f $cGroups)

if ($cGroups -eq 0) {
  Write-Host ""
  Write-Host "  Aucun groupe contenant 'DSIM' trouve sur $($site.Name)." -ForegroundColor Red
  Log "Aucun groupe DSIM trouve. Arret." "ERROR"
  return
}

# ============================================================
# ETAPE 2 : CHOIX INTERACTIF DU GROUPE
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  GROUPES CONTENANT 'DSIM'" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$sortedGroups = $dsimGroups | Sort-Object { $_.name }
$idx = 0
foreach ($g in $sortedGroups) {
  $idx++
  Write-Host ("  {0,3}. {1}" -f $idx, $g.name) -ForegroundColor Yellow
}
Write-Host ""

$choixGroupe = 0
while ($choixGroupe -lt 1 -or $choixGroupe -gt $cGroups) {
  $inputGroupe = Read-Host "  Choisissez un groupe (1-$cGroups)"
  if ($inputGroupe -match "^\d+$") { $choixGroupe = [int]$inputGroupe }
}
$selectedGroup = $sortedGroups[$choixGroupe - 1]
Log ("  Groupe selectionne : {0} (ID: {1})" -f $selectedGroup.name, $selectedGroup.groupId)
Write-Host ""
Write-Host ("  => Groupe selectionne : {0}" -f $selectedGroup.name) -ForegroundColor Green

# ============================================================
# ETAPE 3 : CHOIX DU NIVEAU DE PERMISSION
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  NIVEAU DE PERMISSION" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Lecture                (consulter les pages)" -ForegroundColor Yellow
Write-Host "  2. Ecriture              (creer/modifier pages, blogs, PJ, commentaires)" -ForegroundColor Yellow
Write-Host "  3. Admin                 (administrer l'espace)" -ForegroundColor Yellow
Write-Host "  4. Ecriture+Suppression  (ecriture + supprimer pages, blogs, PJ, commentaires)" -ForegroundColor Yellow
Write-Host ""

$permChoices = @("Lecture", "Ecriture", "Admin", "Ecriture+Suppression")
$choixPerm = 0
while ($choixPerm -lt 1 -or $choixPerm -gt 4) {
  $inputPerm = Read-Host "  Choisissez un niveau (1-4)"
  if ($inputPerm -match "^\d+$") { $choixPerm = [int]$inputPerm }
}
$selectedPermName = $permChoices[$choixPerm - 1]
$selectedPerms = $permissionSets[$selectedPermName]
Log ("  Permission selectionnee : {0} ({1} operations)" -f $selectedPermName, $selectedPerms.Count)
Write-Host ""
Write-Host ("  => Permission : {0}" -f $selectedPermName) -ForegroundColor Green

# ============================================================
# ETAPE 4 : LECTURE DU FICHIER DE SPACE KEYS
# ============================================================
Log "=== ETAPE 4 : Lecture du fichier de space keys ==="
$rawLines = Get-Content -Path $SpaceKeysFile -Encoding UTF8
$spaceKeys = New-Object System.Collections.Generic.List[string]

foreach ($line in $rawLines) {
  $trimmed = $line.Trim()
  if (-not $trimmed) { continue }
  if ($trimmed.StartsWith("#")) { continue }
  if ($trimmed -match "^SpaceKey" -and $spaceKeys.Count -eq 0) { continue }
  $parts = $trimmed -split "[;,`t]"
  $key = $parts[0].Trim().Trim('"')
  if ($key) { $spaceKeys.Add($key) | Out-Null }
}

$cSpaces = ($spaceKeys | Measure-Object).Count
Log ("  {0} space keys chargees depuis {1}" -f $cSpaces, $SpaceKeysFile)

if ($cSpaces -eq 0) {
  Write-Host "  Le fichier ne contient aucune space key valide." -ForegroundColor Red
  return
}
Write-Host ""
Write-Host ("  {0} space keys chargees" -f $cSpaces) -ForegroundColor Green

# ============================================================
# ETAPE 5 : VERIFICATION DES ESPACES (par space key, pas besoin d'ID)
# ============================================================
Log "=== ETAPE 5 : Verification de l'existence des espaces ==="
$spaces = New-Object System.Collections.Generic.List[object]
$cNotFound = 0

foreach ($sk in $spaceKeys) {
  $encodedKey = [System.Uri]::EscapeDataString($sk)
  $url = "{0}/wiki/rest/api/space/{1}" -f $site.BaseUrl, $encodedKey
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $site.Headers

  if (-not $resp.ok) {
    if ($resp.status -eq 404) {
      Log ("  Space {0} non trouve (404)" -f $sk) "WARN"
      $spaces.Add(@{ key=$sk; name=""; status="NON TROUVE"; error="404" }) | Out-Null
    } else {
      Log ("  Erreur verification space {0} : status={1}" -f $sk, $resp.status) "WARN"
      $spaces.Add(@{ key=$sk; name=""; status="ERREUR"; error="status=$($resp.status)" }) | Out-Null
    }
    $cNotFound++
    continue
  }

  $json = $resp.content | ConvertFrom-Json
  $spaces.Add(@{
    key    = $sk
    name   = [string]$json.name
    status = "OK"
    error  = ""
  }) | Out-Null
  Log ("  {0} => {1}" -f $sk, $json.name)
  Start-Sleep -Milliseconds 150
}

$cResolved = ($spaces | Where-Object { $_.status -eq "OK" } | Measure-Object).Count
Log ("  Verifies : {0} / {1}" -f $cResolved, $cSpaces)

if ($cResolved -eq 0) {
  Write-Host "  Aucun espace verifie. Verifiez les space keys." -ForegroundColor Red
  return
}

# ============================================================
# ETAPE 6 : RECAPITULATIF ET APPLICATION
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RECAPITULATIF" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Groupe     : {0}" -f $selectedGroup.name) -ForegroundColor White
Write-Host ("  Permission : {0}" -f $selectedPermName) -ForegroundColor White
Write-Host ("  Espaces    : {0} verifies sur {1}" -f $cResolved, $cSpaces) -ForegroundColor White
Write-Host ("  Mode       : {0}" -f $Mode) -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host ("{0,-6} {1,-15} {2,-40} {3}" -f "#", "SPACE KEY", "NOM", "STATUT") -ForegroundColor DarkGray

$idx = 0
foreach ($sp in $spaces) {
  $idx++
  $color = if ($sp.status -eq "OK") { "Yellow" } else { "Red" }
  $nameDisplay = if ($sp.name) { $sp.name } else { $sp.error }
  Write-Host ("{0,4}. {1,-15} {2,-40} {3}" -f $idx, $sp.key, $nameDisplay, $sp.status) -ForegroundColor $color
}
Write-Host ""

# --- CSV header ---
$csvColumns = @("SpaceKey","SpaceName","Groupe","Permission","Statut","Detail")
$csvHeader = ($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
Write-CsvHeader $csvHeader

if ($Mode -eq "Dry") {
  Write-Host "========================================" -ForegroundColor Green
  Write-Host "  MODE DRY : aucune modification" -ForegroundColor Green
  Write-Host "  Relancez le script et choisissez Execute pour appliquer" -ForegroundColor Green
  Write-Host "========================================" -ForegroundColor Green
  Log "Mode Dry : aucune action effectuee."

  foreach ($sp in $spaces) {
    $row = @($sp.key, $sp.name, $selectedGroup.name, $selectedPermName, $sp.status, $sp.error)
    $csvLine = ($row | ForEach-Object { '"{0}"' -f ([string]$_ -replace '"','""') }) -join ";"
    Write-CsvLine $csvLine
  }
}
else {
  # Confirmation finale
  Write-Host "========================================" -ForegroundColor Red
  Write-Host "  MODE EXECUTE" -ForegroundColor Red
  Write-Host ("  Ajouter '{0}' en '{1}' sur {2} espaces ?" -f $selectedGroup.name, $selectedPermName, $cResolved) -ForegroundColor Red
  Write-Host "========================================" -ForegroundColor Red
  Write-Host ""
  $confirm = Read-Host "  Confirmer ? (OUI pour continuer)"
  if ($confirm -ne "OUI") {
    Write-Host "  Annule." -ForegroundColor Yellow
    Log "Annule par l'utilisateur avant execution."
    return
  }

  Write-Host ""
  Write-Host "  Application en cours..." -ForegroundColor White
  Write-Host ""

  $cOk = 0; $cSkip = 0; $cErr = 0

  foreach ($sp in $spaces) {
    if ($sp.status -ne "OK") {
      Log ("  SKIP {0} : {1}" -f $sp.key, $sp.status)
      $row = @($sp.key, $sp.name, $selectedGroup.name, $selectedPermName, "SKIP", $sp.status)
      $csvLine = ($row | ForEach-Object { '"{0}"' -f ([string]$_ -replace '"','""') }) -join ";"
      Write-CsvLine $csvLine
      $cSkip++
      continue
    }

    Write-Host ("  {0,-15} {1,-35} " -f $sp.key, $sp.name) -NoNewline

    $allOk = $true
    $errDetail = ""
    $encodedSpaceKey = [System.Uri]::EscapeDataString($sp.key)

    foreach ($perm in $selectedPerms) {
      # --- API v1 : POST /wiki/rest/api/space/{key}/permission ---
      $bodyObj = @{
        subject = @{
          type       = "group"
          identifier = $selectedGroup.groupId
        }
        operation = @{
          key    = $perm.operation
          target = $perm.targetType
        }
      }
      $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress

      $url = "{0}/wiki/rest/api/space/{1}/permission" -f $site.BaseUrl, $encodedSpaceKey
      $resp = Invoke-ApiCall -Method "POST" -Url $url -Headers $site.Headers -Body $bodyJson

      if ($resp.ok -or $resp.status -eq 200 -or $resp.status -eq 201) {
        # OK — permission ajoutee
      }
      elseif ($resp.status -eq 400 -and $resp.body -match "already exists") {
        # Permission deja existante — OK
        Log ("    {0}:{1} deja present sur {2}" -f $perm.operation, $perm.targetType, $sp.key)
      }
      elseif ($resp.status -eq 409) {
        # Conflit — permission deja existante (variante)
        Log ("    {0}:{1} deja present sur {2}" -f $perm.operation, $perm.targetType, $sp.key)
      }
      else {
        $allOk = $false
        $errDetail += ("{0}:{1}=ERR{2} " -f $perm.operation, $perm.targetType, $resp.status)
        Log ("    ERREUR {0}:{1} sur {2} : status={3} {4}" -f $perm.operation, $perm.targetType, $sp.key, $resp.status, $resp.error) "ERROR"
      }

      Start-Sleep -Milliseconds 200
    }

    if ($allOk) {
      Write-Host "OK" -ForegroundColor Green
      Log ("  OK : {0} => {1} [{2}]" -f $sp.key, $selectedGroup.name, $selectedPermName)
      $row = @($sp.key, $sp.name, $selectedGroup.name, $selectedPermName, "OK", "")
      $cOk++
    } else {
      Write-Host "ERREUR" -ForegroundColor Red
      $row = @($sp.key, $sp.name, $selectedGroup.name, $selectedPermName, "ERREUR", $errDetail.Trim())
      $cErr++
    }

    $csvLine = ($row | ForEach-Object { '"{0}"' -f ([string]$_ -replace '"','""') }) -join ";"
    Write-CsvLine $csvLine

    Start-Sleep -Milliseconds 300
  }

  Write-Host ""
  Write-Host "========================================" -ForegroundColor Cyan
  Write-Host "  RESULTAT" -ForegroundColor Cyan
  Write-Host "========================================" -ForegroundColor Cyan
  Write-Host ("  OK      : {0}" -f $cOk) -ForegroundColor Green
  Write-Host ("  Skip    : {0}" -f $cSkip) -ForegroundColor DarkGray
  Write-Host ("  Erreurs : {0}" -f $cErr) -ForegroundColor Red
  Log ("Termine : {0} OK, {1} skip, {2} erreurs" -f $cOk, $cSkip, $cErr)
}

# ============================================================
# RESUME
# ============================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Fichier    : {0}" -f $SpaceKeysFile)
Write-Host ("  Groupe     : {0}" -f $selectedGroup.name)
Write-Host ("  Permission : {0}" -f $selectedPermName)
Write-Host ("  Espaces    : {0} verifies / {1} total" -f $cResolved, $cSpaces)
Write-Host ("  Mode       : {0}" -f $Mode)
Write-Host ("  CSV        : {0}" -f $csvFile)
Write-Host ("  LOG        : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan
Log "Termine."