<#
.SYNOPSIS
  Rattachement-Equipe-Jira.ps1
  Rattache automatiquement des tickets Jira à une équipe si le rapporteur
  est membre de cette équipe.

.DESCRIPTION
  Ce script :
  1. Demande via popup la clé du projet et le nom de l'équipe
  2. Récupère les membres de l'équipe Jira via l'API Atlassian Teams
  3. Recherche dans le projet les tickets dont :
     - le rapporteur est membre de l'équipe
     - le champ "Team" n'est pas encore renseigné
  4. Met à jour ces tickets pour y affecter l'équipe.

  MODE DRYRUN :
    Liste les tickets éligibles sans les modifier.
    Le CSV de sortie est écrit ligne par ligne (consultable pendant l'exécution).

  MODE EXECUTE :
    Met réellement à jour le champ "Team" sur les tickets éligibles.

  POPUPS (4 au total) :
    1. Saisie de la clé du projet (ex: CEEP)
    2. Saisie du nom de l'équipe (ex: CEEP - Prestations - Squad 1)
    3. Choix du mode DryRun / Execute / Annuler
    4. Résumé final

  COLONNES DU CSV DE SORTIE :
    IssueKey        => Clé du ticket
    Summary         => Résumé du ticket
    Reporter        => Email du rapporteur
    ReporterName    => Nom du rapporteur
    CurrentTeam     => Équipe actuelle (vide si non renseignée)
    ResultAction    => Résultat de l'action

.NOTES
  Auteur         : Frédéric GUEDJ
  Compatibilité  : PowerShell 5.1+
  Prérequis      :
    - secrets\org-admin.xml    (OrgId + API Key pour l'API Teams)
    - secrets\jira-admin.xml   (Email + API Token pour l'API Jira REST)

  APIs UTILISÉES :
    GET  https://api.atlassian.com/gateway/api/public/teams/v1/org/{orgId}/teams
    GET  https://api.atlassian.com/gateway/api/public/teams/v1/org/{orgId}/teams/{teamId}/members
    POST https://jiradot.atlassian.net/rest/api/3/search
    GET  https://jiradot.atlassian.net/rest/api/3/field
    PUT  https://jiradot.atlassian.net/rest/api/3/issue/{issueKey}

.PARAMETER JiraSite
  URL du site Jira. Par défaut : "jiradot.atlassian.net"

.PARAMETER ThrottleMs
  Délai entre chaque appel API en ms. Par défaut : 200

.EXAMPLE
  .\Rattachement-Equipe-Jira.ps1
  .\Rattachement-Equipe-Jira.ps1 -JiraSite "jiradot.atlassian.net" -ThrottleMs 300
#>

[CmdletBinding()]
param(
  [string] $JiraSite = "jiradot.atlassian.net",
  [int]    $ThrottleMs = 200,
  [int]    $MaxRetries = 10,
  [int]    $PageSize = 50
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null

# -------------------- Paths & log --------------------
$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $ExportsDir ("RattachementEquipe_{0}.log" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# -------------------- Popup générique de saisie --------------------
function Show-InputBox([string]$Title, [string]$Prompt, [string]$Default = "") {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.Size = New-Object System.Drawing.Size(450, 180)
  $form.StartPosition = "CenterScreen"
  $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false

  $label = New-Object System.Windows.Forms.Label
  $label.Location = New-Object System.Drawing.Point(15, 15)
  $label.Size = New-Object System.Drawing.Size(400, 25)
  $label.Text = $Prompt
  $form.Controls.Add($label)

  $textBox = New-Object System.Windows.Forms.TextBox
  $textBox.Location = New-Object System.Drawing.Point(15, 45)
  $textBox.Size = New-Object System.Drawing.Size(400, 25)
  $textBox.Text = $Default
  $form.Controls.Add($textBox)

  $okButton = New-Object System.Windows.Forms.Button
  $okButton.Location = New-Object System.Drawing.Point(250, 90)
  $okButton.Size = New-Object System.Drawing.Size(75, 30)
  $okButton.Text = "OK"
  $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $form.AcceptButton = $okButton
  $form.Controls.Add($okButton)

  $cancelButton = New-Object System.Windows.Forms.Button
  $cancelButton.Location = New-Object System.Drawing.Point(340, 90)
  $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
  $cancelButton.Text = "Annuler"
  $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $form.CancelButton = $cancelButton
  $form.Controls.Add($cancelButton)

  $form.TopMost = $true
  $result = $form.ShowDialog()

  if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    return $textBox.Text.Trim()
  }
  return $null
}

# -------------------- Popup : choix DryRun / Execute / Annuler --------------------
function Choose-Mode {
  $result = [System.Windows.Forms.MessageBox]::Show(
    ("Voulez-vous EXECUTER les mises à jour (Oui) ou faire un DRY RUN (Non) ?`n`n" +
     "Oui     = EXECUTE (modification réelle du champ Team)`n" +
     "Non     = DRY RUN (liste les tickets éligibles sans modifier)`n" +
     "Annuler = Quitter"),
    "Mode d'exécution - Rattachement Equipe",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return "EXECUTE" }
  if ($result -eq [System.Windows.Forms.DialogResult]::No)  { return "DRYRUN" }
  return "CANCEL"
}

# -------------------- POPUP 1 : Clé du projet --------------------
$ProjectKey = Show-InputBox -Title "Projet Jira" -Prompt "Clé du projet Jira (ex: CEEP, TVSP, APER) :" -Default "CEEP"
if ([string]::IsNullOrWhiteSpace($ProjectKey)) { Write-Host "Annulé."; return }
$ProjectKey = $ProjectKey.ToUpper()

# -------------------- POPUP 2 : Nom de l'équipe --------------------
$TeamName = Show-InputBox -Title "Équipe Jira" -Prompt "Nom exact de l'équipe Jira :" -Default "CEEP - Prestations - Squad 1"
if ([string]::IsNullOrWhiteSpace($TeamName)) { Write-Host "Annulé."; return }

Log "Projet=$ProjectKey | Équipe=$TeamName"

# -------------------- Load credentials --------------------
# 1) Org Admin (pour l'API Teams)
$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) { throw "Fichier manquant: $orgCredFile" }
$orgData = Import-Clixml -Path $orgCredFile
$orgId = [string]$orgData.OrgId
$orgApiKey = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password

# 2) Jira Admin (pour l'API Jira REST)
$jiraCredFile = Join-Path $SecretsDir "jira-admin.xml"
if (-not (Test-Path $jiraCredFile)) {
  throw @"
Fichier manquant: $jiraCredFile
Créez-le avec :
  `$cred = @{
    Email = 'votre.email@harmonie-mutuelle.fr'
    ApiTokenSecureString = (Read-Host 'API Token Jira' -AsSecureString)
  }
  `$cred | Export-Clixml -Path '$jiraCredFile'
"@
}
$jiraData = Import-Clixml -Path $jiraCredFile
$jiraEmail = [string]$jiraData.Email
$jiraToken = [System.Net.NetworkCredential]::new("", $jiraData.ApiTokenSecureString).Password
$jiraBasicAuth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${jiraEmail}:${jiraToken}"))

# Headers
$headersOrg = @{
  Authorization  = "Bearer $orgApiKey"
  Accept         = "application/json"
  "Content-Type" = "application/json"
}
$headersJira = @{
  Authorization  = "Basic $jiraBasicAuth"
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
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)

  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{ Method=$Method; Uri=$Url; Headers=$Headers; UseBasicParsing=$true; ErrorAction="Stop" }
      if ($Body) { $params["ContentType"]="application/json"; $params["Body"]=[Text.Encoding]::UTF8.GetBytes($Body) }
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

# ==================== ÉTAPE 1 : Trouver l'équipe ====================
Log "=== ÉTAPE 1 : Recherche de l'équipe '$TeamName' ==="

$teamId = $null
$cursor = $null
$teamFound = $false

do {
  $url = "https://api.atlassian.com/gateway/api/public/teams/v1/org/$orgId/teams?limit=50"
  if ($cursor) { $url += "&cursor=$cursor" }

  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $headersOrg
  if (-not $resp.ok) { throw "Erreur API Teams: status=$($resp.status) $($resp.error) body=$($resp.body)" }

  $json = $resp.content | ConvertFrom-Json
  foreach ($team in $json.entities) {
    if ($team.displayName -eq $TeamName) {
      $teamId = $team.teamId
      $teamFound = $true
      Log "Équipe trouvée: id=$teamId displayName='$($team.displayName)'"
      break
    }
  }

  $cursor = $json.cursor
  if ([string]::IsNullOrWhiteSpace($cursor)) { $cursor = $null }

  Start-Sleep -Milliseconds 100
} while (-not $teamFound -and $cursor)

if (-not $teamFound) { throw "Équipe '$TeamName' introuvable dans l'organisation." }

# ==================== ÉTAPE 2 : Récupérer les membres ====================
Log "=== ÉTAPE 2 : Récupération des membres de l'équipe ==="

$teamMembers = @{}  # accountId => displayName
$cursor = $null

do {
  $url = "https://api.atlassian.com/gateway/api/public/teams/v1/org/$orgId/teams/$teamId/members?limit=50"
  if ($cursor) { $url += "&cursor=$cursor" }

  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $headersOrg
  if (-not $resp.ok) { throw "Erreur API Team Members: status=$($resp.status) $($resp.error)" }

  $json = $resp.content | ConvertFrom-Json
  foreach ($member in $json.entities) {
    $accId = $member.accountId
    $name  = $member.displayName
    if ($accId -and -not $teamMembers.ContainsKey($accId)) {
      $teamMembers[$accId] = $name
    }
  }

  $cursor = $json.cursor
  if ([string]::IsNullOrWhiteSpace($cursor)) { $cursor = $null }

  Start-Sleep -Milliseconds 100
} while ($cursor)

Log "Membres trouvés: $($teamMembers.Count)"
foreach ($m in $teamMembers.GetEnumerator()) {
  Log "  - $($m.Value) ($($m.Key))"
}

if ($teamMembers.Count -eq 0) { throw "Aucun membre trouvé dans l'équipe '$TeamName'." }

# ==================== ÉTAPE 3 : Trouver le champ "Team" ====================
Log "=== ÉTAPE 3 : Identification du champ Team dans Jira ==="

$url = "https://$JiraSite/rest/api/3/field"
$resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $headersJira
if (-not $resp.ok) { throw "Erreur API fields: status=$($resp.status) $($resp.error)" }

$fields = $resp.content | ConvertFrom-Json

# Chercher le champ Team (plusieurs noms possibles)
$teamField = $fields | Where-Object {
  ($_.name -eq "Team") -or
  ($_.schema -and $_.schema.custom -and $_.schema.custom -like "*team*")
} | Select-Object -First 1

if (-not $teamField) {
  $teamField = $fields | Where-Object { $_.name -eq "Équipe" } | Select-Object -First 1
}

if (-not $teamField) { throw "Champ 'Team' introuvable dans les champs Jira." }

$teamFieldId = $teamField.id
Log "Champ Team trouvé: id='$teamFieldId' name='$($teamField.name)'"

# ==================== ÉTAPE 4 : Mode d'exécution ====================
$mode = Choose-Mode
if ($mode -eq "CANCEL") { Log "Annulé." "WARN"; return }
Log "Mode=$mode"

# ==================== ÉTAPE 5 : Recherche des tickets ====================
Log "=== ÉTAPE 5 : Recherche des tickets éligibles ==="

# Construire le JQL avec les accountIds des membres
$memberIdsList = ($teamMembers.Keys | ForEach-Object { "`"$_`"" }) -join ", "
$jql = "project = `"$ProjectKey`" AND reporter IN ($memberIdsList) AND `"$($teamField.name)`" IS EMPTY ORDER BY created DESC"

Log "JQL: $jql"

# Export incrémental
$outCsv = Join-Path $ExportsDir ("RattachementEquipe_{0}_Result_{1}_{2}.csv" -f $ProjectKey, $mode, $ts)
$allColumns = @("IssueKey", "Summary", "Reporter", "ReporterName", "CurrentTeam", "ResultAction")
$headerLine = ($allColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";"
$headerLine | Set-Content -Path $outCsv -Encoding UTF8
Log "Export incrémental -> $outCsv"

# Compteurs
$cTotal = 0; $cUpdated = 0; $cError = 0
$startAt = 0
$hasMore = $true
$totalResults = 0
$startTime = Get-Date

while ($hasMore) {
  $body = @{
    jql        = $jql
    startAt    = $startAt
    maxResults = $PageSize
    fields     = @("summary", "reporter", $teamFieldId)
  } | ConvertTo-Json -Depth 5

  $url = "https://$JiraSite/rest/api/3/search"
  $resp = Invoke-ApiCall -Method "POST" -Url $url -Headers $headersJira -Body $body

  if (-not $resp.ok) {
    Log "Erreur recherche JQL: status=$($resp.status) body=$($resp.body)" "ERROR"
    throw "Erreur recherche: $($resp.error)"
  }

  $searchResult = $resp.content | ConvertFrom-Json
  $totalResults = $searchResult.total
  $issues = $searchResult.issues

  if ($startAt -eq 0) {
    Log "Tickets éligibles trouvés: $totalResults"
    if ($totalResults -eq 0) {
      Log "Aucun ticket à traiter."
      break
    }
  }

  foreach ($issue in $issues) {
    $cTotal++
    $issueKey      = $issue.key
    $summary       = $issue.fields.summary
    $reporterAcct  = $issue.fields.reporter.accountId
    $reporterName  = $issue.fields.reporter.displayName
    $reporterEmail = $issue.fields.reporter.emailAddress
    $currentTeam   = ""

    $teamValue = $issue.fields.PSObject.Properties | Where-Object { $_.Name -eq $teamFieldId } | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue
    if ($teamValue -and $teamValue.name) { $currentTeam = $teamValue.name }
    elseif ($teamValue) { $currentTeam = [string]$teamValue }

    # Estimation temps restant
    $elapsed = (Get-Date) - $startTime
    if ($cTotal -gt 1) {
      $remaining = [Math]::Round(($elapsed.TotalSeconds / ($cTotal - 1)) * ($totalResults - $cTotal) / 60, 1)
      $etaMsg = "~${remaining} min restantes"
    } else {
      $etaMsg = "calcul..."
    }
    Write-Progress -Activity "Rattachement équipe ($mode) - $etaMsg" -Status "$cTotal/$totalResults $issueKey" -PercentComplete ([int](100*$cTotal/[Math]::Max(1,$totalResults)))

    $resultAction = ""

    if ($mode -eq "EXECUTE") {
      # Tentative 1 : format objet { id: teamId }
      $updateBody = @{
        fields = @{
          $teamFieldId = @{ id = $teamId }
        }
      } | ConvertTo-Json -Depth 5

      $updateUrl = "https://$JiraSite/rest/api/3/issue/$issueKey"
      $updateResp = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $headersJira -Body $updateBody

      if ($updateResp.ok -or $updateResp.status -eq 204) {
        $resultAction = "Équipe assignée"
        $cUpdated++
        Log "OK $issueKey => Team='$TeamName'"
      } else {
        # Tentative 2 : format string teamId
        $updateBody2 = @{
          fields = @{
            $teamFieldId = $teamId
          }
        } | ConvertTo-Json -Depth 5

        $updateResp2 = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $headersJira -Body $updateBody2
        if ($updateResp2.ok -or $updateResp2.status -eq 204) {
          $resultAction = "Équipe assignée"
          $cUpdated++
          Log "OK $issueKey => Team='$TeamName' (format alt)"
        } else {
          # Tentative 3 : format name
          $updateBody3 = @{
            fields = @{
              $teamFieldId = @{ name = $TeamName }
            }
          } | ConvertTo-Json -Depth 5

          $updateResp3 = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $headersJira -Body $updateBody3
          if ($updateResp3.ok -or $updateResp3.status -eq 204) {
            $resultAction = "Équipe assignée"
            $cUpdated++
            Log "OK $issueKey => Team='$TeamName' (format name)"
          } else {
            $resultAction = "ERREUR: status=$($updateResp.status) $($updateResp.body)"
            $cError++
            Log "ERREUR $issueKey : s1=$($updateResp.status) s2=$($updateResp2.status) s3=$($updateResp3.status)" "ERROR"
            Log "  Body1: $($updateResp.body)" "ERROR"
          }
        }
      }

      Start-Sleep -Milliseconds $ThrottleMs
    } else {
      $resultAction = "[DRYRUN] A rattacher à '$TeamName'"
    }

    # Écriture incrémentale
    $csvLine = '"{0}";"{1}";"{2}";"{3}";"{4}";"{5}"' -f `
      ($issueKey -replace '"','""'),
      ($summary -replace '"','""'),
      ($(if($reporterEmail){$reporterEmail}else{$reporterAcct}) -replace '"','""'),
      ($reporterName -replace '"','""'),
      ($currentTeam -replace '"','""'),
      ($resultAction -replace '"','""')
    Add-Content -Path $outCsv -Value $csvLine -Encoding UTF8
  }

  $startAt += $issues.Count
  $hasMore = ($startAt -lt $totalResults)

  if ($hasMore) { Start-Sleep -Milliseconds $ThrottleMs }
}

# -------------------- Summary --------------------
Log ("RESUME: Projet={0} Équipe={1} Tickets={2} MisAJour={3} Erreurs={4}" -f $ProjectKey, $TeamName, $cTotal, $cUpdated, $cError)
Log "Terminé."

# -------------------- Popup résumé --------------------
$summaryMsg = "Mode: $mode`n" +
  "Projet: $ProjectKey`n" +
  "Équipe: $TeamName`n" +
  "Membres de l'équipe: $($teamMembers.Count)`n" +
  "`nTickets éligibles: $cTotal`n"

if ($mode -eq "EXECUTE") {
  $summaryMsg += "Mis à jour: $cUpdated`nErreurs: $cError`n"
}

$summaryMsg += "`nExport: $outCsv"

[System.Windows.Forms.MessageBox]::Show($summaryMsg, "Résultat - Rattachement Equipe", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null