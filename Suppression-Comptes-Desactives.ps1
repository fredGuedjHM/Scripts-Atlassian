<#
.SYNOPSIS
  Suppression-Comptes-Desactives.ps1
  Supprime definitivement les comptes Atlassian a partir d un fichier CSV
  prealablement valide (issu du script d analyse).

.DESCRIPTION
  Ce script :
  1. Demande le chemin du fichier CSV d entree (output retravaille du script d analyse)
  2. Charge et affiche les comptes a supprimer
  3. Demande confirmation globale
  4. Supprime au cas par cas (ou en masse avec confirmation renforcee)

  FICHIER D ENTREE :
    Le CSV doit contenir au minimum la colonne "AccountId".
    Colonnes attendues (meme format que l output du script d analyse) :
      AccountId;DisplayName;Email;AccountType;Status;
      Domaines;Groupes;DateCreation;DateInvitation;DerniereActivite;
      DateReference;AncienneteMois
    Le fichier peut avoir ete retravaille (lignes supprimees, commentees, etc.)
    Seules les lignes avec un AccountId valide seront traitees.

  SECURITES :
    - Double confirmation avant toute suppression
    - Mode cas par cas par defaut (O/N/T/Q)
    - Log de chaque action (succes/erreur/ignore)
    - CSV de sortie avec le resultat de chaque suppression
    - Aucune suppression si le fichier est vide ou invalide

  /!\ LA SUPPRESSION EST IRREVERSIBLE /!\
  Un compte supprime ne peut PAS etre restaure.
  Toutes les donnees associees (issues assignees, commentaires, etc.)
  seront orphelines (attribuees a "Former user").

  API :
    POST /users/{accountId}/manage/lifecycle/delete

  CREDENTIALS :
    secrets\org-admin.xml => OrgId + API Key

.NOTES
  Auteur         : Frederic GUEDJ
  Version        : 1.0 — Juillet 2026
  Compatibilite  : PowerShell 5.1+

.EXAMPLE
  .\Suppression-Comptes-Desactives.ps1
  .\Suppression-Comptes-Desactives.ps1 -AutoAll
#>

[CmdletBinding()]
param(
  [string] $FichierCsv = "",

  [switch] $AutoAll,

  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile    = Join-Path $ExportsDir ("Suppression-Comptes_{0}.log" -f $ts)
$csvResultat = Join-Path $ExportsDir ("Suppression-Comptes_Resultat_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
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
# ETAPE 1 : CREDENTIALS
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Red
Write-Host "  SUPPRESSION DEFINITIVE DE COMPTES" -ForegroundColor Red
Write-Host "  /!\ IRREVERSIBLE /!\" -ForegroundColor Red
Write-Host "========================================" -ForegroundColor Red
Write-Host ""
Write-Host "  Ce script supprime definitivement les comptes Atlassian" -ForegroundColor White
Write-Host "  listes dans un fichier CSV valide." -ForegroundColor White
Write-Host ""
Write-Host "  Un compte supprime :" -ForegroundColor Yellow
Write-Host "    - Ne peut PAS etre restaure" -ForegroundColor Yellow
Write-Host "    - Ses issues/commentaires deviennent orphelins" -ForegroundColor Yellow
Write-Host "    - Son historique d activite est perdu" -ForegroundColor Yellow
Write-Host ""

Log "================================================================"
Log "  SUPPRESSION DEFINITIVE DE COMPTES ATLASSIAN"
Log "  /!\ IRREVERSIBLE /!\"
Log "================================================================"

$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) {
  Log "Fichier org-admin.xml introuvable, creation interactive..." "WARN"
  [System.Windows.Forms.MessageBox]::Show(
    ("Le fichier org-admin.xml n'existe pas.`n`nVous allez fournir :`n  1. L'OrgId`n  2. Une API Key d'organisation"),
    "Credentials Organisation", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  $inputOrgId = Read-Host "OrgId (visible dans admin.atlassian.com)"
  $apiKeySecure = Read-Host "API Key d'organisation" -AsSecureString
  @{ OrgId=$inputOrgId; ApiKeySecureString=$apiKeySecure } | Export-Clixml -Path $orgCredFile
  Log "Credentials sauvegardes dans $orgCredFile"
}
$orgData = Import-Clixml -Path $orgCredFile
$orgId = [string]$orgData.OrgId
$orgApiKey = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password
$orgHeaders = @{ Authorization = "Bearer $orgApiKey"; Accept = "application/json" }

Log ("  OrgId : {0}" -f $orgId)

# ============================================================
# ETAPE 2 : CHARGEMENT DU FICHIER CSV
# ============================================================

Log "=== ETAPE 2 : Chargement du fichier CSV ==="

# Demander le fichier si non fourni en parametre
if (-not $FichierCsv -or -not (Test-Path $FichierCsv)) {
  Write-Host "  Selectionnez le fichier CSV des comptes a supprimer :" -ForegroundColor White
  Write-Host ""

  # Proposer un file picker ou saisie manuelle
  $choixMode = Read-Host "  [F]ichier picker ou [S]aisie du chemin ? (F/S)"

  if ($choixMode -match "^[Ff]") {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Title = "Selectionner le CSV des comptes a supprimer"
    $openDialog.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $openDialog.InitialDirectory = $ExportsDir
    $result = $openDialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
      Write-Host "  Annule par l'utilisateur." -ForegroundColor Yellow
      Log "Annule : aucun fichier selectionne."
      exit 0
    }
    $FichierCsv = $openDialog.FileName
  } else {
    $FichierCsv = Read-Host "  Chemin complet du fichier CSV"
  }
}

if (-not (Test-Path $FichierCsv)) {
  Write-Host ("  ERREUR : Fichier introuvable : {0}" -f $FichierCsv) -ForegroundColor Red
  Log ("ERREUR : Fichier introuvable : {0}" -f $FichierCsv) "ERROR"
  exit 1
}

Log ("  Fichier : {0}" -f $FichierCsv)

# Charger le CSV
try {
  $csvContent = Import-Csv -Path $FichierCsv -Delimiter ";" -Encoding UTF8
} catch {
  # Tenter avec le delimiter virgule
  try {
    $csvContent = Import-Csv -Path $FichierCsv -Delimiter "," -Encoding UTF8
  } catch {
    Write-Host ("  ERREUR : Impossible de lire le CSV : {0}" -f $_.Exception.Message) -ForegroundColor Red
    Log ("ERREUR lecture CSV : {0}" -f $_.Exception.Message) "ERROR"
    exit 1
  }
}

# Verifier la presence de la colonne AccountId
$headers = $csvContent | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

if ("AccountId" -notin $headers) {
  Write-Host "  ERREUR : La colonne 'AccountId' est absente du CSV." -ForegroundColor Red
  Write-Host ("  Colonnes trouvees : {0}" -f ($headers -join ", ")) -ForegroundColor DarkGray
  Log "ERREUR : Colonne AccountId absente." "ERROR"
  exit 1
}

# Filtrer les lignes valides (AccountId non vide)
$comptesASupprimer = @()
foreach ($row in $csvContent) {
  $accId = [string]$row.AccountId
  # Ignorer les lignes vides ou commentees (AccountId commencant par #)
  if ([string]::IsNullOrWhiteSpace($accId)) { continue }
  if ($accId.StartsWith("#")) { continue }
  $comptesASupprimer += $row
}

if ($comptesASupprimer.Count -eq 0) {
  Write-Host "  Aucun compte valide trouve dans le fichier." -ForegroundColor Yellow
  Log "Aucun compte valide dans le fichier."
  exit 0
}

Log ("  {0} comptes a traiter" -f $comptesASupprimer.Count)

# ============================================================
# ETAPE 3 : AFFICHAGE ET CONFIRMATION
# ============================================================

Log "=== ETAPE 3 : Verification et confirmation ==="

Write-Host ""
Write-Host "========================================" -ForegroundColor Red
Write-Host ("  {0} COMPTES A SUPPRIMER DEFINITIVEMENT" -f $comptesASupprimer.Count) -ForegroundColor Red
Write-Host "========================================" -ForegroundColor Red
Write-Host ""

$idx = 0
foreach ($c in $comptesASupprimer) {
  $idx++

  $displayName = ""; if ($c.DisplayName) { $displayName = [string]$c.DisplayName }
  $email       = ""; if ($c.Email) { $email = [string]$c.Email }
  $status      = ""; if ($c.Status) { $status = [string]$c.Status }
  $domaines    = ""; if ($c.Domaines) { $domaines = [string]$c.Domaines }
  $lastAct     = ""; if ($c.DerniereActivite) { $lastAct = [string]$c.DerniereActivite }
  $anciennete  = ""; if ($c.AncienneteMois) { $anciennete = [string]$c.AncienneteMois + " mois" }
  $groupes     = ""; if ($c.Groupes) { $groupes = [string]$c.Groupes }

  if ($idx -le 50) {
    Write-Host ("{0,4}. {1}" -f $idx, $displayName) -ForegroundColor Yellow
    Write-Host ("      Email       : {0}" -f $email) -ForegroundColor DarkGray
    Write-Host ("      Status      : {0}" -f $status) -ForegroundColor DarkGray
    Write-Host ("      Domaines    : {0}" -f $domaines) -ForegroundColor DarkGray
    Write-Host ("      Dern. act.  : {0}" -f $lastAct) -ForegroundColor DarkGray
    Write-Host ("      Anciennete  : {0}" -f $anciennete) -ForegroundColor DarkGray
    if ($groupes) {
      Write-Host ("      Groupes     : {0}" -f $groupes) -ForegroundColor DarkCyan
    }
    Write-Host ""
  } elseif ($idx -eq 51) {
    Write-Host ("      ... et {0} autres" -f ($comptesASupprimer.Count - 50)) -ForegroundColor DarkGray
    Write-Host ""
  }
}

# Confirmation globale
Write-Host "========================================" -ForegroundColor Red
Write-Host "  ATTENTION : Cette operation est IRREVERSIBLE." -ForegroundColor Red
Write-Host ("  {0} comptes seront DEFINITIVEMENT supprimes." -f $comptesASupprimer.Count) -ForegroundColor Red
Write-Host "========================================" -ForegroundColor Red
Write-Host ""

$confirm1 = Read-Host "  Confirmer la suppression ? Tapez SUPPRIMER pour continuer"

if ($confirm1 -ne "SUPPRIMER") {
  Write-Host "  Annule. Aucune suppression effectuee." -ForegroundColor Green
  Log "Annule par l'utilisateur (confirmation 1)."
  exit 0
}

$confirm2 = Read-Host "  Derniere chance. Tapez le nombre exact de comptes a supprimer ($($comptesASupprimer.Count))"

if ($confirm2 -ne [string]$comptesASupprimer.Count) {
  Write-Host "  Annule. Le nombre ne correspond pas." -ForegroundColor Green
  Log "Annule par l'utilisateur (confirmation 2 : nombre incorrect)."
  exit 0
}

Log "Double confirmation validee. Debut des suppressions."

# ============================================================
# ETAPE 4 : SUPPRESSION
# ============================================================

Log "=== ETAPE 4 : Suppression des comptes ==="

$cSupprime = 0
$cIgnore   = 0
$cErreur   = 0
$autoAll   = $AutoAll.IsPresent

$resultats = New-Object System.Collections.Generic.List[object]

foreach ($c in $comptesASupprimer) {
  $accId       = [string]$c.AccountId
  $displayName = ""; if ($c.DisplayName) { $displayName = [string]$c.DisplayName }
  $email       = ""; if ($c.Email) { $email = [string]$c.Email }
  $status      = ""; if ($c.Status) { $status = [string]$c.Status }
  $domaines    = ""; if ($c.Domaines) { $domaines = [string]$c.Domaines }

  $doIt = $false

  if ($autoAll) {
    $doIt = $true
    Write-Host ("  [{0}] {1} ({2}) => (auto) Suppression..." -f $accId.Substring(0,8), $displayName, $email) -ForegroundColor DarkYellow
  } else {
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host ("  Nom       : {0}" -f $displayName) -ForegroundColor White
    Write-Host ("  Email     : {0}" -f $email) -ForegroundColor White
    Write-Host ("  Status    : {0}" -f $status) -ForegroundColor White
    Write-Host ("  Domaines  : {0}" -f $domaines) -ForegroundColor White
    Write-Host ("  ID        : {0}" -f $accId) -ForegroundColor DarkGray
    Write-Host ""

    $choice = Read-Host "  SUPPRIMER ce compte ? (O/N/T=tous/Q=quitter)"

    if ($choice -match "^[Qq]") {
      Log "Arret demande par l'utilisateur."
      Write-Host "  Arret." -ForegroundColor Yellow
      # Marquer les restants comme non traites
      $resultats.Add(@{ AccountId=$accId; DisplayName=$displayName; Email=$email; Resultat="Non traite (arret)" }) | Out-Null
      # Ajouter tous les suivants
      $currentIdx = [array]::IndexOf($comptesASupprimer, $c)
      for ($i = $currentIdx + 1; $i -lt $comptesASupprimer.Count; $i++) {
        $remaining = $comptesASupprimer[$i]
        $resultats.Add(@{
          AccountId   = [string]$remaining.AccountId
          DisplayName = if ($remaining.DisplayName) { [string]$remaining.DisplayName } else { "" }
          Email       = if ($remaining.Email) { [string]$remaining.Email } else { "" }
          Resultat    = "Non traite (arret)"
        }) | Out-Null
      }
      break
    }
    elseif ($choice -match "^[Tt]") {
      Write-Host ""
      $confirmAuto = Read-Host "  Confirmer la suppression de TOUS les comptes restants ? Tapez TOUS"
      if ($confirmAuto -eq "TOUS") {
        $autoAll = $true
        $doIt = $true
        Log "Mode auto-all active."
      } else {
        Write-Host "  Auto-all annule, on continue au cas par cas." -ForegroundColor Yellow
        $choice = Read-Host "  Supprimer CE compte ? (O/N)"
        if ($choice -match "^[Oo]") { $doIt = $true }
      }
    }
    elseif ($choice -match "^[Oo]") {
      $doIt = $true
    }
  }

  if ($doIt) {
# APRÈS (avec body JSON obligatoire)
    $deleteUrl = "https://api.atlassian.com/users/$accId/manage/lifecycle/delete"
    $deleteBody = '{"message":"Suppression administrative - compte désactivé depuis plus de 3 mois"}'
    $deleteResp = Invoke-ApiCall -Method "POST" -Url $deleteUrl -Headers $orgHeaders -Body $deleteBody

    if ($deleteResp.ok -or $deleteResp.status -eq 204 -or $deleteResp.status -eq 200) {
      Log ("  SUPPRIME : {0} ({1}) [{2}]" -f $displayName, $email, $accId)
      Write-Host "  => SUPPRIME" -ForegroundColor Green
      $cSupprime++
      $resultats.Add(@{ AccountId=$accId; DisplayName=$displayName; Email=$email; Resultat="Supprime" }) | Out-Null
    } else {
      Log ("  ERREUR suppression {0} : status={1} {2}" -f $displayName, $deleteResp.status, $deleteResp.error) "ERROR"
      Write-Host ("  => ERREUR (status={0})" -f $deleteResp.status) -ForegroundColor Red
      $cErreur++
      $resultats.Add(@{ AccountId=$accId; DisplayName=$displayName; Email=$email; Resultat=("Erreur status={0}" -f $deleteResp.status) }) | Out-Null
    }
  } else {
    Log ("  IGNORE : {0} ({1})" -f $displayName, $email)
    Write-Host "  => Ignore" -ForegroundColor DarkGray
    $cIgnore++
    $resultats.Add(@{ AccountId=$accId; DisplayName=$displayName; Email=$email; Resultat="Ignore" }) | Out-Null
  }

  Start-Sleep -Milliseconds 500
}

# ============================================================
# ETAPE 5 : EXPORT CSV RESULTAT
# ============================================================

Log "=== ETAPE 5 : Export CSV resultat ==="

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('"AccountId";"DisplayName";"Email";"Resultat";"DateTraitement"')

$dateTraitement = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

foreach ($r in $resultats) {
  $line = '"{0}";"{1}";"{2}";"{3}";"{4}"' -f `
    ([string]$r.AccountId -replace '"','""'),
    ([string]$r.DisplayName -replace '"','""'),
    ([string]$r.Email -replace '"','""'),
    ([string]$r.Resultat -replace '"','""'),
    $dateTraitement
  [void]$sb.AppendLine($line)
}

[System.IO.File]::WriteAllText($csvResultat, $sb.ToString(), $utf8Bom)
Log "  CSV resultat -> $csvResultat"

# ============================================================
# RESUME
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME SUPPRESSION" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Comptes dans le fichier         : {0}" -f $comptesASupprimer.Count)
Write-Host ("  SUPPRIMES                       : {0}" -f $cSupprime) -ForegroundColor $(if ($cSupprime -gt 0) { "Green" } else { "DarkGray" })
Write-Host ("  Ignores                         : {0}" -f $cIgnore) -ForegroundColor DarkGray
Write-Host ("  Erreurs                         : {0}" -f $cErreur) -ForegroundColor $(if ($cErreur -gt 0) { "Red" } else { "DarkGray" })
Write-Host ""
Write-Host ("  Fichier source                  : {0}" -f $FichierCsv)
Write-Host ("  CSV resultat                    : {0}" -f $csvResultat)
Write-Host ("  LOG                             : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

if ($cErreur -gt 0) {
  Write-Host ""
  Write-Host "  /!\ Des erreurs sont survenues. Verifiez le log et le CSV resultat." -ForegroundColor Red
}

Log ("Termine : {0} supprimes, {1} ignores, {2} erreurs" -f $cSupprime, $cIgnore, $cErreur)