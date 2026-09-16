<#
.SYNOPSIS
  Desactivation-Suppression-Users-v2.ps1
  Supprime les comptes désactivés depuis plus de 3 mois, restreint à certains domaines.

.DESCRIPTION
  APPROCHE HYBRIDE (optimale) :
  1. Liste les users via GET /admin/v1/orgs/{orgId}/users (paginé, ~50/page)
  2. Pour chaque page, filtre par domaine email autorisé
  3. Les comptes actifs sont écrits immédiatement dans le CSV (aucun appel API)
  4. Les comptes non actifs sont vérifiés via Events API (?q=email)
     pour trouver la date de désactivation
  5. Si désactivé depuis > 90 jours => éligible à la suppression

  Le CSV est écrit AU FIL DE L'EAU : consultable dès le début de l'exécution.

  DOMAINES AUTORISÉS :
    harmonie-mutuelle.fr, prestataire.sihm.fr, mutex.fr,
    mutex-exterieur.fr, chorum.fr

  MODE DRYRUN : liste les comptes éligibles sans supprimer.
  MODE EXECUTE : supprime réellement (double confirmation).

  COLONNES DU CSV DE SORTIE :
    AccountId           => ID du compte Atlassian
    Email               => Email du compte
    Name                => Nom affiché
    Domain              => Domaine email
    AccountStatus       => Statut du compte
    DateDesactivation   => Date de désactivation (DD/MM/YYYY HH:mm)
    JoursDepuis         => Nombre de jours depuis la désactivation
    Eligible            => Oui/Non
    ResultAction        => Résultat de l'action

.NOTES
  Auteur         : Frédéric GUEDJ
  Compatibilité  : PowerShell 5.1+
  Prérequis      : secrets\org-admin.xml (OrgId + API Key)

  APIs UTILISÉES :
    GET  /admin/v1/orgs/{orgId}/users              (liste paginée)
    GET  /admin/v1/orgs/{orgId}/events?q={email}    (date de désactivation)
    POST /users/{accountId}/manage/lifecycle/delete  (suppression)

.EXAMPLE
  .\Desactivation-Suppression-Users-v2.ps1
  .\Desactivation-Suppression-Users-v2.ps1 -SeuilJours 120
#>

[CmdletBinding()]
param(
  [int] $ThrottleMs = 400,
  [int] $MaxRetries = 10,
  [int] $SeuilJours = 90
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# -------------------- Domaines autorisés --------------------
$domainesAutorises = @(
  "harmonie-mutuelle.fr",
  "prestataire.sihm.fr",
  "mutex.fr",
  "mutex-exterieur.fr",
  "chorum.fr"
)

# -------------------- Date de référence --------------------
$dateRef = (Get-Date).Date
$dateSeuil = $dateRef.AddDays(-$SeuilJours)

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("SuppressionDesactives_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# -------------------- Popup : choix DryRun / Execute / Annuler --------------------
function Choose-Mode {
  $result = [System.Windows.Forms.MessageBox]::Show(
    ("Suppression des comptes désactivés depuis plus de $SeuilJours jours.`n" +
     "Domaines : $($domainesAutorises -join ', ')`n`n" +
     "Oui     = EXECUTE (suppressions REELLES)`n" +
     "Non     = DRY RUN (liste les comptes éligibles sans supprimer)`n" +
     "Annuler = Quitter"),
    "Mode d'exécution - Suppression comptes désactivés",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return "EXECUTE" }
  if ($result -eq [System.Windows.Forms.DialogResult]::No)  { return "DRYRUN" }
  return "CANCEL"
}

# -------------------- Load credentials --------------------
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }

$data = Import-Clixml -Path $orgCredFile
$orgId = [string]$data.OrgId
$apiKeyPlain = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password
$headers = @{
  Authorization  = "Bearer $apiKeyPlain"
  Accept         = "application/json"
  "Content-Type" = "application/json"
}

# -------------------- Network --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# -------------------- HTTP helper --------------------
function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [string]$Body = $null)

  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{ Method=$Method; Uri=$Url; Headers=$script:headers; UseBasicParsing=$true; ErrorAction="Stop" }
      if ($Body) { $params["ContentType"]="application/json"; $params["Body"]=$Body }
      $resp = Invoke-WebRequest @params
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$resp.Content }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $reader.ReadToEnd(); $reader.Close()
      } catch {}

      if ($attempt -gt $MaxRetries) {
        return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
      }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6, $attempt)))
        Log "Retry $Method status=$status in ${sleepSec}s ($attempt/$MaxRetries)" "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }
      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# -------------------- API : Events par email (date désactivation) --------------------
function Get-DeactivationDate([string]$Email, [string]$AccountId) {
  $encodedEmail = [System.Uri]::EscapeDataString($Email)
  $url = "https://api.atlassian.com/admin/v1/orgs/$script:orgId/events?q=$encodedEmail&limit=20"
  $resp = Invoke-ApiCall -Method "GET" -Url $url

  if (-not $resp.ok) {
    Log "  Events API erreur pour $Email : status=$($resp.status)" "WARN"
    return $null
  }

  try {
    $json = $resp.content | ConvertFrom-Json
    if (-not $json.data) { return $null }

    $deactActions = @("managed_account_deactivated", "user_disabled_org_access", "managed_account_deleted")
    foreach ($evt in $json.data) {
      $evtAction = [string]$evt.attributes.action
      if ($evtAction -in $deactActions) {
        $matchUser = $false
        foreach ($ctx in $evt.attributes.context) {
          if ($ctx.type -eq "users" -and $ctx.id -eq $AccountId) { $matchUser = $true; break }
        }
        if (-not $matchUser) {
          foreach ($ctn in $evt.attributes.container) {
            if ($ctn.type -eq "users" -and $ctn.id -eq $AccountId) { $matchUser = $true; break }
          }
        }
        if ($matchUser) {
          $eventTime = [datetime]::Parse($evt.attributes.time)
          return @{ date=$eventTime; action=$evtAction }
        }
      }
    }
  } catch {
    Log "  Erreur parsing events $Email : $($_.Exception.Message)" "WARN"
  }
  return $null
}

# -------------------- API : Delete --------------------
function Delete-User([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/delete"
  return Invoke-ApiCall -Method "POST" -Url $url -Body '{"message":"Suppression automatique - compte désactivé depuis plus de 3 mois"}'
}

# -------------------- Extraction du domaine email --------------------
function Get-EmailDomain([string]$email) {
  if ($email -match "@(.+)$") { return $matches[1].ToLower() }
  return ""
}

# ==================== MAIN ====================

$mode = Choose-Mode
if ($mode -eq "CANCEL") { Log "Annulé." "WARN"; return }

# Double confirmation en mode EXECUTE
if ($mode -eq "EXECUTE") {
  $confirm = [System.Windows.Forms.MessageBox]::Show(
    ("ATTENTION : Vous allez SUPPRIMER DEFINITIVEMENT les comptes`n" +
     "désactivés depuis plus de $SeuilJours jours.`n`n" +
     "Domaines concernés :`n" +
     ($domainesAutorises | ForEach-Object { "  - $_" } | Out-String) +
     "Cette action est IRRÉVERSIBLE.`n`nConfirmer la suppression ?"),
    "CONFIRMATION SUPPRESSION",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Stop
  )
  if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
    Log "Suppression annulée par l'utilisateur." "WARN"
    return
  }
}

Log "Mode=$mode | Seuil=$SeuilJours jours | Date seuil=$($dateSeuil.ToString('dd/MM/yyyy'))"
Log "Domaines autorisés: $($domainesAutorises -join ', ')"

# -------------------- Export incrémental : header --------------------
$outCsv = Join-Path $ExportsDir ("SuppressionDesactives_Result_{0}_{1}.csv" -f $mode, $ts)
$allColumns = @("AccountId", "Email", "Name", "Domain", "AccountStatus", "DateDesactivation", "JoursDepuis", "Eligible", "ResultAction")
$headerLine = ($allColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
$headerLine | Set-Content -Path $outCsv -Encoding UTF8
Log "Export incrémental -> $outCsv (consultable pendant l'exécution)"

# Compteurs
$cTotalLus = 0; $cDomainOk = 0; $cActif = 0; $cNonActif = 0
$cEligible = 0; $cNonEligible = 0; $cDateInconnue = 0
$cDeleted = 0; $cError = 0

$pageNum = 0
$nextUrl = "https://api.atlassian.com/admin/v1/orgs/$orgId/users?limit=100"
$startTime = Get-Date

Log "=== Parcours des comptes de l'organisation ==="

while ($nextUrl) {
  $pageNum++

  $resp = Invoke-ApiCall -Method "GET" -Url $nextUrl
  if (-not $resp.ok) { throw "Erreur API users page $pageNum : status=$($resp.status) $($resp.error)" }

  $json = $resp.content | ConvertFrom-Json
  $usersInPage = $json.data
  $cTotalLus += $usersInPage.Count

  $domainOkInPage = 0
  $nonActifInPage = 0

  foreach ($user in $usersInPage) {
    $accId  = [string]$user.account_id
    $email  = [string]$user.email
    $name   = [string]$user.name
    $status = [string]$user.account_status
    $domain = Get-EmailDomain $email

    # Filtre domaine
    if ($domain -notin $domainesAutorises) { continue }

    $cDomainOk++
    $domainOkInPage++

    $dateDesactivation = ""
    $joursDepuis       = ""
    $eligible          = "Non"
    $resultAction      = ""

    # ---- ACTIF : aucun appel API, écriture directe ----
    if ($status -eq "active") {
      $cActif++
      $resultAction = "Compte actif"

      $out = [ordered]@{
        AccountId=$accId; Email=$email; Name=$name; Domain=$domain
        AccountStatus=$status; DateDesactivation=""; JoursDepuis=""
        Eligible="Non"; ResultAction=$resultAction
      }
      $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
      Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8
      continue
    }

    # ---- NON ACTIF : chercher la date de désactivation ----
    $cNonActif++
    $nonActifInPage++

    Write-Progress -Activity "Vérification ($mode)" -Status "Page $pageNum | $email" -PercentComplete -1

    $evtInfo = Get-DeactivationDate -Email $email -AccountId $accId

    if ($evtInfo) {
      $deactDate = $evtInfo.date.ToLocalTime()
      $dateDesactivation = $deactDate.ToString("dd/MM/yyyy HH:mm")
      $jours = [int]($dateRef - $deactDate.Date).TotalDays
      $joursDepuis = $jours

      if ($jours -ge $SeuilJours) {
        $eligible = "Oui"
        $cEligible++

        if ($mode -eq "EXECUTE") {
          Log "  Suppression $email (désactivé le $dateDesactivation, $jours jours)..."
          $delResp = Delete-User -AccountId $accId
          if ($delResp.ok -or $delResp.status -eq 204) {
            $resultAction = "Supprimé (désactivé depuis $jours jours)"
            $cDeleted++
            Log "  OK supprimé $email"
          } else {
            $resultAction = "ERREUR SUPPRESSION: status=$($delResp.status)"
            $cError++
            Log "  ERREUR suppression $email : status=$($delResp.status) body=$($delResp.body)" "ERROR"
          }
        } else {
          $resultAction = "[DRYRUN] A supprimer (désactivé depuis $jours jours)"
          Log "  ELIGIBLE: $email désactivé le $dateDesactivation ($jours jours)"
        }
      } else {
        $cNonEligible++
        $resultAction = "Non supprimé ($jours jours < $SeuilJours)"
        Log "  $email => $jours jours < $SeuilJours, non éligible"
      }
    } else {
      $cDateInconnue++
      $resultAction = "Non supprimé (date désactivation inconnue)"
      Log "  $email => désactivé (status=$status) mais date inconnue"
    }

    # Écriture incrémentale
    $out = [ordered]@{
      AccountId=$accId; Email=$email; Name=$name; Domain=$domain
      AccountStatus=$status; DateDesactivation=$dateDesactivation
      JoursDepuis=$joursDepuis; Eligible=$eligible; ResultAction=$resultAction
    }
    $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join ";"
    Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8

    Start-Sleep -Milliseconds $ThrottleMs
  }

  # Estimation temps
  $elapsed = (Get-Date) - $startTime
  $elapsedMin = [Math]::Round($elapsed.TotalMinutes, 1)

  Log ("  Page {0}: {1} users lus, {2} domaine OK ({3} non actifs) | Total: {4} lus, {5} domaine OK, {6} actifs, {7} non actifs | {8} min écoulées" -f `
    $pageNum, $usersInPage.Count, $domainOkInPage, $nonActifInPage, `
    $cTotalLus, $cDomainOk, $cActif, $cNonActif, $elapsedMin)

  # Pagination
  $nextUrl = $null
  if ($json.links -and $json.links.next) {
    $nextUrl = [string]$json.links.next
  }

  Start-Sleep -Milliseconds 200
}

# -------------------- Summary --------------------
Log "============================================"
Log "RESUME"
Log "============================================"
Log "Domaines: $($domainesAutorises -join ', ')"
Log "Seuil: $SeuilJours jours (avant le $($dateSeuil.ToString('dd/MM/yyyy')))"
Log "Total comptes lus: $cTotalLus"
Log "  Domaines autorisés: $cDomainOk"
Log "    Actifs: $cActif"
Log "    Non actifs vérifiés: $cNonActif"
Log "      Éligibles (>=$SeuilJours jours): $cEligible"
Log "      Non éligibles (<$SeuilJours jours): $cNonEligible"
Log "      Date inconnue: $cDateInconnue"
if ($mode -eq "EXECUTE") {
  Log "      Supprimés: $cDeleted"
  Log "      Erreurs: $cError"
}
Log "============================================"
$totalMin = [Math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
Log "Durée totale: $totalMin minutes"
Log "Terminé."

# -------------------- Popup résumé --------------------
$summaryMsg = "Mode: $mode`n" +
  "Seuil: $SeuilJours jours`n" +
  "Domaines: $($domainesAutorises -join ', ')`n" +
  "`nTotal comptes lus: $cTotalLus`n" +
  "Domaines OK: $cDomainOk`n" +
  "  Actifs: $cActif`n" +
  "  Non actifs: $cNonActif`n" +
  "`nÉligibles: $cEligible`n" +
  "Non éligibles: $cNonEligible`n" +
  "Date inconnue: $cDateInconnue`n"

if ($mode -eq "EXECUTE") {
  $summaryMsg += "`nSupprimés: $cDeleted`nErreurs: $cError`n"
}

$summaryMsg += "`nDurée: $totalMin min`nExport: $outCsv"

[System.Windows.Forms.MessageBox]::Show($summaryMsg, "Résultat - Suppression comptes désactivés", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null