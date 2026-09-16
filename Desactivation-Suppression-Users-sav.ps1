<#
.SYNOPSIS
  Desactivation-Suppression-Users.ps1
  Désactivation ou suppression en masse de comptes utilisateurs Atlassian Cloud.

.DESCRIPTION
  Ce script lit un fichier CSV contenant une liste d'utilisateurs avec une colonne ACTION
  qui détermine le traitement à appliquer à chaque compte :

    - "A desactiver"                                 => Désactivation de l'utilisateur
                                                        (POST /users/{id}/manage/lifecycle/disable)

    - "A supprimer"                                  => Suppression de l'utilisateur
                                                        (POST /users/{id}/manage/lifecycle/delete)

    - "?"                                            => Aucune action (skip)
    - "A garder"                                     => Aucune action (skip)
    - "A supprimer si descativer avant le 01/01/26"  => Aucune action (skip)
    - Toute autre valeur                             => Aucune action (skip)

  Les deux endpoints utilisent l'API User Management et fonctionnent avec
  l'Org Admin API Key (pas besoin d'OAuth).

  RÉSULTATS POSSIBLES (colonne ResultAction) :
    "Désactivé"                     => disable OK (204)
    "Supprimé"                      => delete OK (204)
    "Non modifié"                   => ACTION = skip
    "[DRYRUN] A désactiver"         => mode simulation
    "[DRYRUN] A supprimer"          => mode simulation
    "ERREUR: ..."                   => erreur

.NOTES
  Nom du script  : Desactivation-Suppression-Users.ps1
  Auteur         : Frédéric GUEDJ
  Compatibilité  : PowerShell 5.1+
  Prérequis      : secrets\org-admin.xml (créé via Save-AtlassianOrgAdminKey.ps1)

  COLONNES ATTENDUES DANS LE CSV D'ENTRÉE :
    AccountId ; Email ; EmailDomain ; Name ; AccountStatus ; AccessBillable ;
    ACTION ; ProductAccessCount ; Apps ; OrgLastActive ; OrgLastActiveUtc ;
    MaxAppLastActiveUtc ; EffectiveLastActiveUtc

  AUTHENTIFICATION :
    Le script utilise le fichier secrets\org-admin.xml.
    Ce fichier contient l'OrgId et l'API Key chiffrée.

  APIs UTILISÉES (User Management API, auth Org Admin API Key) :

    - Désactivation (disable user) :
        POST /users/{accountId}/manage/lifecycle/disable
        Réponse attendue : 204 No Content

    - Suppression (delete user) :
        POST /users/{accountId}/manage/lifecycle/delete
        Réponse attendue : 204 No Content

  POPUPS (3 au total) :
    1. Sélection du fichier CSV (OpenFileDialog)
    2. Choix du mode d'exécution (MessageBox Oui/Non/Annuler)
       => Oui     = EXECUTE
       => Non     = DRYRUN (simulation)
       => Annuler = le script s'arrête
    3. Résumé final (MessageBox informative)

  FICHIERS PRODUITS (dans .\exports\) :
    - DesactivationSuppression_Result_{DRYRUN|EXECUTE}_{timestamp}.csv
    - DesactivationSuppression_{timestamp}.log

  WORKFLOW RECOMMANDÉ :
    1. Préparer le CSV avec la colonne ACTION renseignée.
    2. Lancer le script, sélectionner le CSV, choisir DRYRUN.
    3. Vérifier le fichier résultat (colonne ResultAction).
    4. Si tout est conforme, relancer le script et choisir EXECUTE.

.EXAMPLE
  .\Desactivation-Suppression-Users.ps1

.EXAMPLE
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
  $dlg.Title = $title
  $dlg.Filter = "CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  $null = $dlg.ShowDialog()
  return $dlg.FileName
}

# -------------------- Popup 2 : choix DryRun / Execute / Annuler --------------------
function Choose-Mode {
  $result = [System.Windows.Forms.MessageBox]::Show(
    ("Voulez-vous EXECUTER les actions (Oui) ou faire un DRY RUN simulation (Non) ?`n`n" +
     "Oui     = EXECUTE (désactivations et suppressions REELLES)`n" +
     "Non     = DRY RUN (simulation, aucune modification)`n" +
     "Annuler = Quitter sans rien faire"),
    "Mode d'exécution",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return "EXECUTE" }
  if ($result -eq [System.Windows.Forms.DialogResult]::No) { return "DRYRUN" }
  return "CANCEL"
}

# -------------------- Load credentials --------------------
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }

$data = Import-Clixml -Path $orgCredFile
$orgId = [string]$data.OrgId
$apiKeyPlain = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password
$headers = @{
  Authorization = "Bearer $apiKeyPlain"
  Accept = "application/json"
  "Content-Type" = "application/json"
}

# -------------------- Network (TLS + proxy) --------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# -------------------- HTTP helper --------------------
function Invoke-ApiCall {
  param(
    [string]$Method,
    [string]$Url,
    [string]$Body = $null
  )

  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{
        Method = $Method
        Uri = $Url
        Headers = $script:headers
        UseBasicParsing = $true
        ErrorAction = "Stop"
      }
      if ($Body) {
        $params["ContentType"] = "application/json"
        $params["Body"] = $Body
      }
      $resp = Invoke-WebRequest @params
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$resp.Content }
    } catch {
      $status = 0
      $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $reader.ReadToEnd()
        $reader.Close()
      } catch {}

      if ($attempt -gt $MaxRetries) {
        return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
      }

      $retryable = ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0)
      if ($retryable) {
        $sleepSec = [Math]::Min(90, [Math]::Pow(2, [Math]::Min(6, $attempt)))
        Log "Retry $Method status=$status in ${sleepSec}s (attempt $attempt/$MaxRetries)" "WARN"
        Start-Sleep -Seconds $sleepSec
        continue
      }

      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# -------------------- API calls (User Management API) --------------------

# Désactivation = disable
# POST /users/{accountId}/manage/lifecycle/disable
# Réponse attendue : 204 No Content
function Disable-User([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/disable"
  return Invoke-ApiCall -Method "POST" -Url $url -Body '{"message":"Désactivation du compte utilisateur"}'
}

# Suppression = delete
# POST /users/{accountId}/manage/lifecycle/delete
# Réponse attendue : 204 No Content
function Delete-User([string]$AccountId) {
  $url = "https://api.atlassian.com/users/$AccountId/manage/lifecycle/delete"
  return Invoke-ApiCall -Method "POST" -Url $url -Body '{"message":"Suppression du compte utilisateur"}'
}

# -------------------- ACTION classification --------------------
function Classify-Action([string]$action) {
  $a = $action.Trim()
  if ($a -eq "A desactiver") { return "DESACTIVER" }
  if ($a -eq "A supprimer")  { return "SUPPRIMER" }
  return "SKIP"
}

# ==================== MAIN ====================

# Popup 1 : sélection du CSV
$csvPath = Pick-File "Sélectionner DesactivationUsers.csv"
if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path $csvPath)) {
  throw "Aucun fichier CSV sélectionné."
}
Log "CSV=$csvPath"

# Popup 2 : choix du mode
$mode = Choose-Mode
if ($mode -eq "CANCEL") {
  Log "Annulé par l'utilisateur." "WARN"
  return
}
Log "Mode=$mode"

# Lecture du CSV
$rawHeader = Get-Content -Path $csvPath -TotalCount 1 -Encoding UTF8
$delim = ','; if ($rawHeader -match ';') { $delim = ';' }

$rows = Import-Csv -Path $csvPath -Delimiter $delim
if (-not $rows -or $rows.Count -eq 0) { throw "CSV vide." }

Log ("Rows loaded={0} delimiter=[{1}]" -f $rows.Count, $delim)

# Preview
$previewDesac = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "DESACTIVER" }).Count
$previewSuppr = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "SUPPRIMER" }).Count
$previewSkip  = @($rows | Where-Object { (Classify-Action $_.ACTION) -eq "SKIP" }).Count

Log ("Preview: A désactiver={0} A supprimer={1} Skip={2}" -f $previewDesac, $previewSuppr, $previewSkip)

# Compteurs
$countDesactiver = 0
$countSupprimer = 0
$countSkip = 0
$countOkDesac = 0
$countOkSuppr = 0
$countError = 0

# Traitement ligne par ligne
$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($r in $rows) {
  $i++
  $accId = ([string]$r.AccountId).Trim()
  $email = [string]$r.Email
  $action = [string]$r.ACTION
  $class = Classify-Action $action

  Write-Progress -Activity "Traitement users ($mode)" -Status "$i/$($rows.Count) $email" -PercentComplete ([int](100*$i/$rows.Count))

  $resultAction = "Non modifié"

  if ($class -eq "SKIP") {
    $countSkip++
    $resultAction = "Non modifié"
  }
  elseif ($class -eq "DESACTIVER") {
    $countDesactiver++
    if ($mode -eq "EXECUTE") {
      if ([string]::IsNullOrWhiteSpace($accId)) {
        $resultAction = "ERREUR: AccountId vide"
        $countError++
      } else {
        Log "Désactivation $email ($accId)..."
        $resp = Disable-User -AccountId $accId

        if ($resp.ok -or $resp.status -eq 204) {
          $resultAction = "Désactivé"
          $countOkDesac++
          Log "OK désactivé $email ($accId) status=$($resp.status)"
        } else {
          $resultAction = "ERREUR DESACTIVATION: status=$($resp.status) $($resp.error)"
          $countError++
          Log "ERREUR désactivation $email ($accId): status=$($resp.status) $($resp.error) body=$($resp.body)" "ERROR"
        }

        Start-Sleep -Milliseconds $ThrottleMs
      }
    } else {
      $resultAction = "[DRYRUN] A désactiver"
    }
  }
  elseif ($class -eq "SUPPRIMER") {
    $countSupprimer++
    if ($mode -eq "EXECUTE") {
      if ([string]::IsNullOrWhiteSpace($accId)) {
        $resultAction = "ERREUR: AccountId vide"
        $countError++
      } else {
        Log "Suppression $email ($accId)..."
        $resp = Delete-User -AccountId $accId

        if ($resp.ok -or $resp.status -eq 204) {
          $resultAction = "Supprimé"
          $countOkSuppr++
          Log "OK supprimé $email ($accId) status=$($resp.status)"
        } else {
          $resultAction = "ERREUR SUPPRESSION: status=$($resp.status) $($resp.error)"
          $countError++
          Log "ERREUR suppression $email ($accId): status=$($resp.status) $($resp.error) body=$($resp.body)" "ERROR"
        }

        Start-Sleep -Milliseconds $ThrottleMs
      }
    } else {
      $resultAction = "[DRYRUN] A supprimer"
    }
  }

  # Construire la ligne de sortie
  $out = [ordered]@{}
  foreach ($p in $r.PSObject.Properties) {
    $out[$p.Name] = $p.Value
  }
  $out["ResultAction"] = $resultAction

  $results.Add([pscustomobject]$out) | Out-Null
}

# -------------------- Export --------------------
$outCsv = Join-Path $ExportsDir ("DesactivationSuppression_Result_{0}_{1}.csv" -f $mode, $ts)
$results | Export-Csv -Path $outCsv -Delimiter $delim -NoTypeInformation -Encoding UTF8
Log "Export -> $outCsv"

# -------------------- Summary --------------------
Log ("RESUME: Total={0} Skip={1} ADesactiver={2} ASupprimer={3}" -f $rows.Count, $countSkip, $countDesactiver, $countSupprimer)
if ($mode -eq "EXECUTE") {
  Log ("EXECUTE: Désactivés={0} Supprimés={1} Erreurs={2}" -f $countOkDesac, $countOkSuppr, $countError)
}
Log "Terminé."

# -------------------- Popup 3 : résumé final --------------------
$summaryMsg = "Mode: $mode`n" +
  "Total lignes: $($rows.Count)`n" +
  "Skip (Non modifié): $countSkip`n" +
  "A désactiver: $countDesactiver`n" +
  "A supprimer: $countSupprimer`n"

if ($mode -eq "EXECUTE") {
  $summaryMsg += "`nDésactivés OK: $countOkDesac"
  $summaryMsg += "`nSupprimés OK: $countOkSuppr"
  $summaryMsg += "`nErreurs: $countError"
}

$summaryMsg += "`n`nExport: $outCsv"

[System.Windows.Forms.MessageBox]::Show(
  $summaryMsg,
  "Résultat - Desactivation-Suppression-Users",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null