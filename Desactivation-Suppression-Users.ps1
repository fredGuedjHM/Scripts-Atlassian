<#
.SYNOPSIS
  Desactivation-Suppression-Users.ps1
  Désactivation ou suppression en masse de comptes utilisateurs Atlassian Cloud.

.DESCRIPTION
  Ce script lit un fichier CSV contenant une liste d'utilisateurs avec une colonne ACTION.

  MODE DRYRUN :
    Pour chaque utilisateur du CSV :
    - "A garder" => aucun appel API, ligne écrite immédiatement (ultra rapide)
    - "?" => 2 appels API (profil + events) pour vérifier le compte douteux
    - Autres actions => 2 appels API (profil + events)
    Aucune modification n'est effectuée. Permet de vérifier l'état réel des comptes.
    Le CSV de sortie est écrit ligne par ligne : consultable pendant l'exécution.

  MODE EXECUTE :
    Exécute réellement les actions (disable / delete) selon la colonne ACTION.

  ACTIONS RECONNUES (colonne ACTION du CSV) :
    "A desactiver"                                 => Désactivation
    "A supprimer"                                  => Suppression
    "A supprimer si descativer avant le 01/01/26"  => Suppression conditionnelle
    "?"                                            => Aucune action mais vérification complète
    "A garder"                                     => Aucune action, aucun appel API
    Toute autre valeur                             => Aucune action, aucun appel API

  COLONNES AJOUTÉES AU CSV DE SORTIE :
    ResultAction        => Résultat de l'action (ou simulation)
    AccountStatusAPI    => Statut actuel remonté par l'API profil
    DateDesactivation   => Date de la dernière désactivation (DD/MM/YYYY HH:mm)

.NOTES
  Auteur         : Frédéric GUEDJ
  Compatibilité  : PowerShell 5.1+
  Prérequis      : secrets\org-admin.xml

  APIs UTILISÉES :
    GET  /users/{accountId}/manage/profile
    GET  /admin/v1/orgs/{orgId}/events?q={email}
    POST /users/{accountId}/manage/lifecycle/disable
    POST /users/{accountId}/manage/lifecycle/delete

.EXAMPLE
  .\Desactivation-Suppression-Users.ps1
  .\Desactivation-Suppression-Users.ps1 -ThrottleMs 200
#>

[CmdletBinding()]
param(
  [int] $ThrottleMs = 400,
  [int] $MaxRetries = 10
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("DesactivationSuppression_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# -------------------- Popup 1 : sélection du fichier CSV --------------------
function Pick-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title  = $title
  $dlg.Filter = "CSV (*.csv)|*.csv|Tous (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  $null = $dlg.ShowDialog()
  return $dlg.FileName
}

# -------------------- Popup 2 : choix DryRun / Execute / Annuler --------------------
function Choose-Mode {
  $result = [System.Windows.Forms.MessageBox]::Show(
    ("Voulez-vous EXECUTER les actions (Oui) ou faire un DRY RUN (Non) ?`n`n" +
     "Oui     = EXECUTE (désactivations et suppressions REELLES)`n" +
     "Non     = DRY RUN (vérification statut + date des comptes concernés)`n" +
     "Annuler = Quitter"),
    "Mode d'exécution",
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

# -------------------- Date limite conditionnelle --------------------
$cutoffDate = [datetime]::new(2026, 1, 1)

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

# -------------------- API : Profil --------------------
function Get-UserProfile([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/profile"
  return Invoke-ApiCall -Method "GET" -Url $url
}

# -------------------- API : Events audit (date désactivation) --------------------
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
          Log "  Event trouvé: action=$evtAction date=$($eventTime.ToLocalTime().ToString('dd/MM/yyyy HH:mm')) pour $Email"
          return @{ date=$eventTime; action=$evtAction }
        }
      }
    }
  } catch {
    Log "  Erreur parsing events $Email : $($_.Exception.Message)" "WARN"
  }
  return $null
}

# -------------------- API : Disable / Delete --------------------
function Disable-User([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/disable"
  return Invoke-ApiCall -Method "POST" -Url $url -Body '{"message":"Désactivation du compte utilisateur"}'
}

function Delete-User([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/delete"
  return Invoke-ApiCall -Method "POST" -Url $url -Body '{"message":"Suppression du compte utilisateur"}'
}

# -------------------- Extraction profil --------------------
function Get-ProfileInfo([string]$AccountId, [string]$Email) {
  $result = @{ accountStatusAPI=""; isDisabled=$false }

  $resp = Get-UserProfile -AccountId $AccountId
  if (-not $resp.ok) {
    $result.accountStatusAPI = "ERREUR_API:$($resp.status)"
    return $result
  }

  try {
    $json = $resp.content | ConvertFrom-Json
    $status = $null
    if ($json.account -and $json.account.account_status) { $status = [string]$json.account.account_status }
    elseif ($json.account_status) { $status = [string]$json.account_status }
    $result.accountStatusAPI = $status

    if ($status -and ($status -in @("inactive","suspended","closed","disabled"))) {
      $result.isDisabled = $true
    }
  } catch {
    Log "  Erreur parsing profil $Email : $($_.Exception.Message)" "ERROR"
  }
  return $result
}

# -------------------- ACTION classification --------------------
# SKIP      = "A garder" et toute autre valeur non reconnue => 0 appel API
# DOUTEUX   = "?" => 2 appels API (profil + events)
# DESACTIVER / SUPPRIMER / SUPPRIMER_CONDITIONNEL => 2 appels API
function Classify-Action([string]$action) {
  $a = $action.Trim()
  if ($a -eq "A desactiver") { return "DESACTIVER" }
  if ($a -eq "A supprimer")  { return "SUPPRIMER" }
  if ($a -eq "A supprimer si descativer avant le 01/01/26") { return "SUPPRIMER_CONDITIONNEL" }
  if ($a -eq "?")            { return "DOUTEUX" }
  return "SKIP"
}

# ==================== MAIN ====================

$csvPath = Pick-File "Sélectionner DesactivationUsers.csv"
if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path $csvPath)) { throw "Aucun fichier CSV sélectionné." }
Log "CSV=$csvPath"

$mode = Choose-Mode
if ($mode -eq "CANCEL") { Log "Annulé." "WARN"; return }
Log "Mode=$mode"

$rawHeader = Get-Content -Path $csvPath -TotalCount 1 -Encoding UTF8
$delim = ','; if ($rawHeader -match ';') { $delim = ';' }
$rows = Import-Csv -Path $csvPath -Delimiter $delim
if (-not $rows -or $rows.Count -eq 0) { throw "CSV vide." }
Log ("Rows={0} delim=[{1}]" -f $rows.Count, $delim)

$previewDesac   = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "DESACTIVER" }).Count
$previewSuppr   = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "SUPPRIMER" }).Count
$previewCond    = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "SUPPRIMER_CONDITIONNEL" }).Count
$previewDouteux = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "DOUTEUX" }).Count
$previewSkip    = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "SKIP" }).Count
Log ("Preview: Désactiver={0} Supprimer={1} Conditionnel={2} Douteux={3} Skip={4}" -f $previewDesac, $previewSuppr, $previewCond, $previewDouteux, $previewSkip)

# Estimation du temps
$apiRows = $previewDesac + $previewSuppr + $previewCond + $previewDouteux
$estMinutes = [Math]::Round($apiRows * 2 * ($ThrottleMs / 1000 + 0.5) / 60, 1)
Log ("Estimation: {0} lignes nécessitant des appels API, ~{1} min" -f $apiRows, $estMinutes)

# -------------------- Export incrémental : header --------------------
$outCsv = Join-Path $ExportsDir ("DesactivationSuppression_Result_{0}_{1}.csv" -f $mode, $ts)
$sampleProps = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
$allColumns  = $sampleProps + @("ResultAction", "AccountStatusAPI", "DateDesactivation")
$headerLine  = ($allColumns | ForEach-Object { '"{0}"' -f $_ }) -join $delim
$headerLine | Set-Content -Path $outCsv -Encoding UTF8
Log "Export incrémental -> $outCsv (consultable pendant l'exécution)"

# Compteurs
$cSkip=0; $cDesac=0; $cSuppr=0; $cCond=0; $cDouteux=0
$cOkDesac=0; $cOkSuppr=0; $cOkCondSuppr=0; $cCondNonElig=0; $cError=0
$cProfileOk=0; $cProfileErr=0

$i = 0
$iApi = 0
$startTime = Get-Date

foreach ($r in $rows) {
  $i++
  $accId  = ([string]$r.AccountId).Trim()
  $email  = [string]$r.Email
  $action = [string]$r.ACTION
  $class  = Classify-Action $action

  # Estimation du temps restant (basée sur les lignes API uniquement)
  $elapsed = (Get-Date) - $startTime
  if ($iApi -gt 0) {
    $avgPerApiItem = $elapsed.TotalSeconds / $iApi
    $apiRemaining  = $apiRows - $iApi
    $remaining = [Math]::Round($avgPerApiItem * [Math]::Max(0, $apiRemaining) / 60, 1)
    $etaMsg = "~${remaining} min restantes"
  } else {
    $etaMsg = "~$estMinutes min estimées"
  }

  Write-Progress -Activity "Traitement ($mode) - $etaMsg" -Status "$i/$($rows.Count) $email" -PercentComplete ([int](100*$i/$rows.Count))

  $resultAction      = "Non modifié"
  $accountStatusAPI  = ""
  $dateDesactivation = ""
  $info    = $null
  $evtInfo = $null

  # ==================== SKIP : aucun appel API ====================
  if ($class -eq "SKIP") {
    $cSkip++
    $resultAction = if ($mode -eq "DRYRUN") { "[DRYRUN] Non modifié" } else { "Non modifié" }
  }
  # ==================== DRYRUN : vérification ====================
  elseif ($mode -eq "DRYRUN") {

    if (-not [string]::IsNullOrWhiteSpace($accId)) {
      # 1) Profil => statut actuel
      $info = Get-ProfileInfo -AccountId $accId -Email $email
      $accountStatusAPI = $info.accountStatusAPI
      Log "  $email => status=$accountStatusAPI"

      if ($info.accountStatusAPI -like "ERREUR*") { $cProfileErr++ } else { $cProfileOk++ }

      # 2) Events => date de désactivation
      $evtInfo = Get-DeactivationDate -Email $email -AccountId $accId
      if ($evtInfo) {
        $dateDesactivation = $evtInfo.date.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
        Log "  $email => désactivé le $dateDesactivation (action=$($evtInfo.action))"
      }

      $iApi++
      Start-Sleep -Milliseconds $ThrottleMs
    } else {
      $accountStatusAPI = "ERREUR: AccountId vide"
      $cProfileErr++
    }

    switch ($class) {
      "DOUTEUX"    { $cDouteux++; $resultAction = "[DRYRUN] Non modifié (douteux)" }
      "DESACTIVER" { $cDesac++;   $resultAction = "[DRYRUN] A désactiver" }
      "SUPPRIMER"  { $cSuppr++;   $resultAction = "[DRYRUN] A supprimer" }
      "SUPPRIMER_CONDITIONNEL" {
        $cCond++
        if (-not $info -or -not $info.isDisabled) {
          $resultAction = "[DRYRUN] Non modifié (compte actif)"
          $cCondNonElig++
        }
        elseif (-not $evtInfo) {
          $resultAction = "[DRYRUN] Non modifié (désactivé, date inconnue)"
          $cCondNonElig++
        }
        elseif ($evtInfo.date -lt $cutoffDate) {
          $resultAction = "[DRYRUN] A supprimer (désactivé le $dateDesactivation, avant 01/01/26)"
        }
        else {
          $resultAction = "[DRYRUN] Non modifié (désactivé le $dateDesactivation, après 01/01/26)"
          $cCondNonElig++
        }
      }
    }
  }
  # ==================== EXECUTE ====================
  else {
    switch ($class) {
      "DOUTEUX" {
        $cDouteux++
        $resultAction = "Non modifié (douteux)"
      }
      "DESACTIVER" {
        $cDesac++
        if ([string]::IsNullOrWhiteSpace($accId)) {
          $resultAction = "ERREUR: AccountId vide"; $cError++
        } else {
          Log "Désactivation $email ($accId)..."
          $resp = Disable-User -AccountId $accId
          if ($resp.ok -or $resp.status -eq 204) {
            $resultAction = "Désactivé"; $cOkDesac++
            Log "OK désactivé $email status=$($resp.status)"
          } else {
            $resultAction = "ERREUR DESACTIVATION: status=$($resp.status) $($resp.error)"; $cError++
            Log "ERREUR désactivation $email : status=$($resp.status) body=$($resp.body)" "ERROR"
          }
          $iApi++
          Start-Sleep -Milliseconds $ThrottleMs
        }
      }
      "SUPPRIMER" {
        $cSuppr++
        if ([string]::IsNullOrWhiteSpace($accId)) {
          $resultAction = "ERREUR: AccountId vide"; $cError++
        } else {
          Log "Suppression $email ($accId)..."
          $resp = Delete-User -AccountId $accId
          if ($resp.ok -or $resp.status -eq 204) {
            $resultAction = "Supprimé"; $cOkSuppr++
            Log "OK supprimé $email status=$($resp.status)"
          } else {
            $resultAction = "ERREUR SUPPRESSION: status=$($resp.status) $($resp.error)"; $cError++
            Log "ERREUR suppression $email : status=$($resp.status) body=$($resp.body)" "ERROR"
          }
          $iApi++
          Start-Sleep -Milliseconds $ThrottleMs
        }
      }
      "SUPPRIMER_CONDITIONNEL" {
        $cCond++
        if ([string]::IsNullOrWhiteSpace($accId)) {
          $resultAction = "ERREUR: AccountId vide"; $cError++
        } else {
          Log "Vérification conditionnelle $email ($accId)..."
          $info = Get-ProfileInfo -AccountId $accId -Email $email
          $accountStatusAPI = $info.accountStatusAPI
          $evtInfo = Get-DeactivationDate -Email $email -AccountId $accId
          if ($evtInfo) { $dateDesactivation = $evtInfo.date.ToLocalTime().ToString("dd/MM/yyyy HH:mm") }

          if (-not $info.isDisabled) {
            $resultAction = "Non modifié (compte actif)"; $cCondNonElig++
          }
          elseif (-not $evtInfo) {
            $resultAction = "Non modifié (désactivé, date inconnue)"; $cCondNonElig++
          }
          elseif ($evtInfo.date -lt $cutoffDate) {
            Log "  Suppression conditionnelle $email (désactivé le $dateDesactivation)..."
            $resp = Delete-User -AccountId $accId
            if ($resp.ok -or $resp.status -eq 204) {
              $resultAction = "Supprimé (désactivé le $dateDesactivation, avant 01/01/26)"; $cOkCondSuppr++
            } else {
              $resultAction = "ERREUR SUPPRESSION COND: status=$($resp.status) $($resp.error)"; $cError++
            }
          }
          else {
            $resultAction = "Non modifié (désactivé le $dateDesactivation, après 01/01/26)"; $cCondNonElig++
          }
          $iApi++
          Start-Sleep -Milliseconds $ThrottleMs
        }
      }
    }
  }

  # -------------------- Écriture incrémentale de la ligne --------------------
  $out = [ordered]@{}
  foreach ($p in $r.PSObject.Properties) { $out[$p.Name] = $p.Value }
  $out["ResultAction"]      = $resultAction
  $out["AccountStatusAPI"]  = $accountStatusAPI
  $out["DateDesactivation"] = $dateDesactivation

  $csvLine = ($allColumns | ForEach-Object { '"{0}"' -f ([string]$out[$_] -replace '"','""') }) -join $delim
  Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8
}

# -------------------- Summary --------------------
Log ("RESUME: Total={0} Skip={1} Douteux={2} Desactiver={3} Supprimer={4} Conditionnel={5}" -f $rows.Count, $cSkip, $cDouteux, $cDesac, $cSuppr, $cCond)
if ($mode -eq "EXECUTE") {
  Log ("EXECUTE: Désactivés={0} Supprimés={1} SupprCond={2} NonElig={3} Erreurs={4}" -f $cOkDesac, $cOkSuppr, $cOkCondSuppr, $cCondNonElig, $cError)
} else {
  Log ("DRYRUN: Profils OK={0} Erreurs={1} CondNonElig={2}" -f $cProfileOk, $cProfileErr, $cCondNonElig)
}
Log "Terminé."

# -------------------- Popup 3 --------------------
$summaryMsg = "Mode: $mode`nTotal: $($rows.Count)`nSkip (A garder): $cSkip`nDouteux (?): $cDouteux`nA désactiver: $cDesac`nA supprimer: $cSuppr`nConditionnel: $cCond`n"
if ($mode -eq "EXECUTE") {
  $summaryMsg += "`nDésactivés OK: $cOkDesac`nSupprimés OK: $cOkSuppr`nSuppr. cond. OK: $cOkCondSuppr`nNon éligibles: $cCondNonElig`nErreurs: $cError"
} else {
  $summaryMsg += "`nProfils OK: $cProfileOk`nProfils erreur: $cProfileErr`nCond. non éligibles: $cCondNonElig"
}
$summaryMsg += "`n`nExport: $outCsv"

[System.Windows.Forms.MessageBox]::Show($summaryMsg, "Résultat", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null