<#
mvp-reporting_v2.ps1 - Lanceur principal v2.1
Corrections v2.1:
  - [bool] au lieu de [switch] avec $true (PSScriptAnalyzer)
  - Initialize-Directory au lieu de Ensure-Dir (verbe approuvé)
  - Assets : endpoint gateway v1 + résolution attributs via dictionnaire objectTypeAttributes
  - XLSX : ImportExcel (sans COM) ou COM en STA runspace avec timeout 120s
  - Encoding : UTF-8 BOM

Lancement:
    .\mvp-reporting_v2.ps1                    # popup mois + popup cache
    .\mvp-reporting_v2.ps1 -SkipCachePrompt   # worklogs+assets rafraîchis, reste cache
    .\mvp-reporting_v2.ps1 -From 2026-04-01 -To 2026-04-30
#>

[CmdletBinding()]
param(
    [datetime]$From, [datetime]$To,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [switch]$ResetTempoToken,
    [bool]$UseSystemProxy = $true,
    [bool]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [switch]$SkipAssets,
    [bool]$CreateXlsx = $true,
    [switch]$SkipCachePrompt
)

# ======================================================================
# DOSSIERS
# ======================================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$sheetsDir  = Join-Path $scriptDir "sheets"
$secretsDir = Join-Path $scriptDir "secrets"
$cacheDir   = Join-Path $scriptDir "cache"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Initialize-Directory([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Initialize-Directory $secretsDir; Initialize-Directory $cacheDir; Initialize-Directory $logsDir
Initialize-Directory $exportsDir; Initialize-Directory $sheetsDir

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile  = Join-Path $logsDir ("mvp-reporting_{0}.log" -f $runStamp)

# ======================================================================
# CHARGEMENT MODULES
# ======================================================================
. (Join-Path $sheetsDir "Common.ps1")
. (Join-Path $sheetsDir "Sheet1-WorklogsIssues.ps1")
. (Join-Path $sheetsDir "Sheet2-SaisiTempsTempo.ps1")
. (Join-Path $sheetsDir "Sheet3-SaisiTempsAsset.ps1")
. (Join-Path $sheetsDir "Sheet4-Anomalies.ps1")
. (Join-Path $sheetsDir "Sheet5-WorklogsSansDate.ps1")

Write-Log "=== DÉBUT EXÉCUTION MVP-REPORTING v2.1 ===" "INFO"
Write-Log "ScriptDir=$scriptDir | LogFile=$logFile" "DEBUG"

# ======================================================================
# PROXY
# ======================================================================
Initialize-Proxy -UseSystemProxy:$UseSystemProxy -ProxyUrl $ProxyUrl

# ======================================================================
# POPUPS
# ======================================================================
function Show-MonthPickerDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MVP-Reporting v2 - Sélection du mois"
    $form.Size = New-Object System.Drawing.Size(350, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Sélectionnez le mois à extraire :"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($label)

    $comboMois = New-Object System.Windows.Forms.ComboBox
    $comboMois.DropDownStyle = "DropDownList"
    $comboMois.Location = New-Object System.Drawing.Point(20, 50)
    $comboMois.Size = New-Object System.Drawing.Size(140, 25)
    @("Janvier","Février","Mars","Avril","Mai","Juin",
      "Juillet","Août","Septembre","Octobre","Novembre","Décembre") |
        ForEach-Object { $comboMois.Items.Add($_) | Out-Null }
    $today = Get-Date
    $comboMois.SelectedIndex = $today.AddMonths(-1).Month - 1
    $form.Controls.Add($comboMois)

    $comboAnnee = New-Object System.Windows.Forms.ComboBox
    $comboAnnee.DropDownStyle = "DropDownList"
    $comboAnnee.Location = New-Object System.Drawing.Point(180, 50)
    $comboAnnee.Size = New-Object System.Drawing.Size(80, 25)
    @(($today.Year - 1), $today.Year, ($today.Year + 1)) |
        ForEach-Object { $comboAnnee.Items.Add($_) | Out-Null }
    $comboAnnee.SelectedItem = $today.AddMonths(-1).Year
    $form.Controls.Add($comboAnnee)

    $labelInfo = New-Object System.Windows.Forms.Label
    $labelInfo.Text = "(Par défaut : mois précédent)"
    $labelInfo.ForeColor = [System.Drawing.Color]::Gray
    $labelInfo.Location = New-Object System.Drawing.Point(20, 85)
    $labelInfo.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($labelInfo)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(80, 115)
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(180, 115)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Info "Extraction annulée."; exit 0
    }
    $moisChoisi = $comboMois.SelectedIndex + 1
    $anneeChoisie = [int]$comboAnnee.SelectedItem
    $form.Dispose()
    $debut = Get-Date -Year $anneeChoisie -Month $moisChoisi -Day 1
    return [pscustomobject]@{
        From = $debut.Date
        To   = $debut.AddMonths(1).AddDays(-1).Date
    }
}

function Show-CacheRefreshDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MVP-Reporting v2 - Gestion du cache"
    $form.Size = New-Object System.Drawing.Size(420, 350)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $lt = New-Object System.Windows.Forms.Label
    $lt.Text = "Sélectionnez les données à rafraîchir :"
    $lt.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lt.Location = New-Object System.Drawing.Point(20, 15)
    $lt.Size = New-Object System.Drawing.Size(370, 20)
    $form.Controls.Add($lt)

    $li = New-Object System.Windows.Forms.Label
    $li.Text = "(Cochés = rechargés depuis les API. Décochés = cache local)"
    $li.ForeColor = [System.Drawing.Color]::Gray
    $li.Location = New-Object System.Drawing.Point(20, 38)
    $li.Size = New-Object System.Drawing.Size(370, 30)
    $form.Controls.Add($li)

    $y = 75
    $chkW = New-Object System.Windows.Forms.CheckBox
    $chkW.Text = "Worklogs Tempo (saisie de temps)"
    $chkW.Checked = $true
    $chkW.Location = New-Object System.Drawing.Point(30, $y)
    $chkW.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkW); $y += 28

    $chkA = New-Object System.Windows.Forms.CheckBox
    $chkA.Text = "Fiches Assets (Employés / Prestataires)"
    $chkA.Checked = $true
    $chkA.Location = New-Object System.Drawing.Point(30, $y)
    $chkA.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkA); $y += 28

    $sep = New-Object System.Windows.Forms.Label
    $sep.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $sep.Location = New-Object System.Drawing.Point(20, ($y + 5))
    $sep.Size = New-Object System.Drawing.Size(360, 2)
    $form.Controls.Add($sep); $y += 18

    $chkT = New-Object System.Windows.Forms.CheckBox
    $chkT.Text = "Équipes Tempo + Workload Schemes"
    $chkT.Checked = $false
    $chkT.Location = New-Object System.Drawing.Point(30, $y)
    $chkT.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkT); $y += 28

    $chkI = New-Object System.Windows.Forms.CheckBox
    $chkI.Text = "Issues Jira (tickets référencés)"
    $chkI.Checked = $false
    $chkI.Location = New-Object System.Drawing.Point(30, $y)
    $chkI.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkI); $y += 28

    $chkG = New-Object System.Windows.Forms.CheckBox
    $chkG.Text = "Groupes Jira (classification Direction)"
    $chkG.Checked = $false
    $chkG.Location = New-Object System.Drawing.Point(30, $y)
    $chkG.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkG); $y += 28

    $chkC = New-Object System.Windows.Forms.CheckBox
    $chkC.Text = "Comptes Tempo (catégories, clients)"
    $chkC.Checked = $false
    $chkC.Location = New-Object System.Drawing.Point(30, $y)
    $chkC.Size = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkC); $y += 40

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Lancer"
    $btnOK.Location = New-Object System.Drawing.Point(100, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(210, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Info "Extraction annulée."; exit 0
    }
    $choices = [pscustomobject]@{
        RefreshWorklogs      = $chkW.Checked
        RefreshAssets        = $chkA.Checked
        RefreshTeams         = $chkT.Checked
        RefreshIssues        = $chkI.Checked
        RefreshGroups        = $chkG.Checked
        RefreshTempoAccounts = $chkC.Checked
    }
    $form.Dispose()
    return $choices
}
# ======================================================================
# PÉRIODE
# ======================================================================
if (-not $From -or -not $To) {
    $periode = Show-MonthPickerDialog
    $From = $periode.From; $To = $periode.To
}
if ($From -gt $To) { throw "From > To ($From > $To)" }
$fromStr = $From.ToString("yyyy-MM-dd")
$toStr   = $To.ToString("yyyy-MM-dd")
Write-Info "Période: $($From.ToString('dd/MM/yyyy')) -> $($To.ToString('dd/MM/yyyy'))"

$feries = Get-JoursFeries -Annee $From.Year
if ($To.Year -ne $From.Year) { $feries += Get-JoursFeries -Annee $To.Year }
$joursOuvres = 0; $d = $From
while ($d -le $To) {
    if ($d.DayOfWeek -ne [DayOfWeek]::Saturday -and
        $d.DayOfWeek -ne [DayOfWeek]::Sunday -and
        $feries -notcontains $d.Date) { $joursOuvres++ }
    $d = $d.AddDays(1)
}
$heuresAttendue = $joursOuvres * 8
Write-Info "Jours ouvrés (hors fériés): $joursOuvres ($heuresAttendue h)"

$weekBuckets = Get-MonthWeekBuckets -MonthStart $From -MonthEnd $To
for ($wi = 0; $wi -lt 5; $wi++) {
    $b = $weekBuckets[$wi]
    if ($null -ne $b.Start) {
        Write-Info "Semaine $($wi+1): $($b.Start.ToString('dd/MM')) -> $($b.End.ToString('dd/MM'))"
    } else {
        Write-Info "Semaine $($wi+1): (vide)"
    }
}

# ======================================================================
# CACHE
# ======================================================================
if (-not $SkipCachePrompt) {
    $cacheChoices = Show-CacheRefreshDialog
} else {
    $cacheChoices = [pscustomobject]@{
        RefreshWorklogs=$true; RefreshAssets=$true; RefreshTeams=$false
        RefreshIssues=$false; RefreshGroups=$false; RefreshTempoAccounts=$false
    }
}

$UseCache_Worklogs      = -not $cacheChoices.RefreshWorklogs
$UseCache_Assets        = -not $cacheChoices.RefreshAssets
$UseCache_Teams         = -not $cacheChoices.RefreshTeams
$UseCache_Issues        = -not $cacheChoices.RefreshIssues
$UseCache_Groups        = -not $cacheChoices.RefreshGroups
$UseCache_TempoAccounts = -not $cacheChoices.RefreshTempoAccounts

$refreshList = @(); $cachedList = @()
if ($cacheChoices.RefreshWorklogs)      { $refreshList += "Worklogs" }      else { $cachedList += "Worklogs" }
if ($cacheChoices.RefreshAssets)        { $refreshList += "Assets" }        else { $cachedList += "Assets" }
if ($cacheChoices.RefreshTeams)         { $refreshList += "Teams+WS" }      else { $cachedList += "Teams+WS" }
if ($cacheChoices.RefreshIssues)        { $refreshList += "Issues" }        else { $cachedList += "Issues" }
if ($cacheChoices.RefreshGroups)        { $refreshList += "Groupes" }       else { $cachedList += "Groupes" }
if ($cacheChoices.RefreshTempoAccounts) { $refreshList += "Comptes Tempo" } else { $cachedList += "Comptes Tempo" }
if ($refreshList.Count -gt 0) { Write-Info "Rafraîchi: $($refreshList -join ', ')" }
else { Write-Info "Cache: tout en cache" }
if ($cachedList.Count -gt 0) { Write-Info "Conservé: $($cachedList -join ', ')" }

$csvSubDir = Join-Path $exportsDir ("csv-MVP-Reporting_{0}_{1}_{2}" -f $fromStr, $toStr, $runStamp)
Initialize-Directory $csvSubDir

# ======================================================================
# CREDENTIALS
# ======================================================================
$jiraCredFile   = Join-Path $secretsDir "jira-jiradot.cred.xml"
$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"

if (-not (Test-Path $jiraCredFile)) {
    throw "Fichier Jira creds introuvable: $jiraCredFile`nLance d'abord Save-JiraCredential.ps1."
}
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password
$jiraHeaders = New-BasicAuthHeader -User $jiraEmail -Pass $jiraToken
Write-Info "Jira: $jiraBaseUrl (user=$jiraEmail)"

function Get-TempoToken {
    param([switch]$Reset)
    if ($Reset -and (Test-Path $tempoTokenFile)) { Remove-Item -Force $tempoTokenFile }
    if (Test-Path $tempoTokenFile) {
        $obj = Import-Clixml -Path $tempoTokenFile
        if ($obj -and $obj.Token) { return (ConvertFrom-SecureStringToPlain -Secure $obj.Token) }
    }
    Write-Host "`nTempo API token requis (sauvegardé DPAPI dans: $tempoTokenFile)"
    $sec = Read-Host -AsSecureString -Prompt "Colle ton Tempo API token (masqué)"
    if (-not $sec) { throw "Tempo token vide." }
    [pscustomobject]@{ Token = $sec } | Export-Clixml -Path $tempoTokenFile
    Write-Info "Tempo token sauvegardé."
    return (ConvertFrom-SecureStringToPlain -Secure $sec)
}
$tempoToken   = Get-TempoToken -Reset:$ResetTempoToken
$tempoHeaders = @{ Authorization = "Bearer $tempoToken"; Accept = "application/json" }

# ======================================================================
# FICHIERS CACHE
# ======================================================================
$worklogCacheFile   = Join-Path $cacheDir "worklogs.json"
$teamsCacheFile     = Join-Path $cacheDir "tempo_teams.json"
$wsCacheFile        = Join-Path $cacheDir "tempo_workload_schemes.json"
$wsDaysCacheFile    = Join-Path $cacheDir "tempo_workload_days.json"
$assetsCacheFile    = Join-Path $cacheDir "saisi_temps_asset.json"
$issuesCacheFile    = Join-Path $cacheDir "issues.json"
$tempoAcctCache     = Join-Path $cacheDir "tempo_accounts_metadata.json"
$groupsCachePath    = Join-Path $cacheDir "audit_user_groups.json"
$jiraUsersCachePath = Join-Path $cacheDir "audit_jira_users.json"

# ======================================================================
# CHARGEMENT : WORKLOGS
# ======================================================================
$worklogs = $null
if ($UseCache_Worklogs) { $worklogs = Load-Json -Path $worklogCacheFile }
if (-not $worklogs) {
    $worklogs = @()
    $url = "https://api.eu.tempo.io/4/worklogs?from=$fromStr&to=$toStr"
    while ($true) {
        $resp = Invoke-ApiGet -Url $url -Headers $tempoHeaders
        if ($resp.results) { $worklogs += $resp.results }
        if ($resp.metadata -and $resp.metadata.next) { $url = $resp.metadata.next }
        else { break }
    }
    Save-Json -Path $worklogCacheFile -Object $worklogs
    Write-Info "Worklogs récupérés: $($worklogs.Count)"
} else { Write-Info "Worklogs cache: $($worklogs.Count)" }

# ======================================================================
# CHARGEMENT : TEAMS
# ======================================================================
$teamsData = $null
if ($UseCache_Teams) { $teamsData = Load-Json -Path $teamsCacheFile }
if (-not $teamsData) {
    $teamsData = @()
    $url = "https://api.tempo.io/4/teams"
    while ($true) {
        $resp = Invoke-ApiGet -Url $url -Headers $tempoHeaders
        foreach ($t in ($resp.results | ForEach-Object { $_ })) {
            $mResp = Invoke-ApiGet -Url "$($t.self)/members" -Headers $tempoHeaders
            $memberIds = @(); $memberDates = @{}
            foreach ($m in ($mResp.results | ForEach-Object { $_ })) {
                if ($m.member -and $m.member.accountId) {
                    $mid = [string]$m.member.accountId
                    $memberIds += $mid
                    $memberDates[$mid] = if ($m.to) { [string]$m.to } else { "" }
                }
            }
            $teamsData += [pscustomobject]@{
                Nom                 = (Fix-DoubleUtf8 $t.name)
                Programme           = $(if ($t.program) { Fix-DoubleUtf8 $t.program.name } else { "Non défini" })
                "Nombre de Membres" = [string]$memberIds.Count
                "Membres ID"        = ($memberIds -join ", ")
                "MemberDates"       = $memberDates
            }
        }
        if ($resp.metadata -and $resp.metadata.next) { $url = $resp.metadata.next }
        else { break }
    }
    Save-Json -Path $teamsCacheFile -Object $teamsData
    Write-Info "Teams récupérées: $($teamsData.Count)"
} else { Write-Info "Teams cache: $($teamsData.Count)" }

# ======================================================================
# CHARGEMENT : WORKLOAD SCHEMES (noms)
# ======================================================================
$workloadSchemeMap = @{}
if ($UseCache_Teams) {
    $cachedWS = Load-Json -Path $wsCacheFile
    if ($cachedWS) {
        foreach ($p in $cachedWS.PSObject.Properties) {
            $workloadSchemeMap[$p.Name] = [string]$p.Value
        }
        Write-Info "Workload schemes cache: $($workloadSchemeMap.Count) users"
    }
}
if ($workloadSchemeMap.Count -eq 0) {
    $wsUrl = "https://api.tempo.io/4/workload-schemes"
    while ($true) {
        $resp = Invoke-ApiGet -Url $wsUrl -Headers $tempoHeaders
        foreach ($ws in ($resp.results | ForEach-Object { $_ })) {
            $schemeName = Fix-DoubleUtf8 ([string]$ws.name)
            $schemeId   = [string]$ws.id
            try {
                $mResp = Invoke-ApiGet -Url "https://api.tempo.io/4/workload-schemes/$schemeId/members" -Headers $tempoHeaders
                foreach ($m in ($mResp.results | ForEach-Object { $_ })) {
                    if ($m.member -and $m.member.accountId) {
                        $workloadSchemeMap[[string]$m.member.accountId] = $schemeName
                    }
                    if ($m.accountId) {
                        $workloadSchemeMap[[string]$m.accountId] = $schemeName
                    }
                }
            } catch { Write-Warn "Workload scheme $schemeId : $($_.Exception.Message)" }
        }
        if ($resp.metadata -and $resp.metadata.next) { $wsUrl = $resp.metadata.next }
        else { break }
    }
    Save-Json -Path $wsCacheFile -Object $workloadSchemeMap
    Write-Info "Workload schemes: $($workloadSchemeMap.Count) users"
}

# ======================================================================
# CHARGEMENT : WORKLOAD SCHEME DAYS (v2 — détail jours par personne)
# ======================================================================
$userWorkloadDays = Load-WorkloadSchemeDays `
    -TempoHeaders $tempoHeaders `
    -CacheFilePath $wsDaysCacheFile `
    -UseCache $UseCache_Teams

# ======================================================================
# CHARGEMENT : ASSETS (v2.1 — gateway + dictionnaire objectTypeAttributes)
# ======================================================================
$assetUsers = @()
if (-not $SkipAssets) {
    if ($UseCache_Assets) {
        $cached = Load-Json -Path $assetsCacheFile
        if ($cached) { $assetUsers = $cached; Write-Info "Assets cache: $($assetUsers.Count)" }
    }
    if ($assetUsers.Count -eq 0) {
        $assetsBaseUrl = "$jiraBaseUrl/gateway/api/jsm/assets/workspace/$AssetsWorkspaceId/v1/object/aql"
        $aql = '(objectType = Employé OR objectType = Prestataire)'
        $startAt = 0; $maxResults = 50; $isLast = $false

        $wantedMap = @{
            "Compte Jira"    = "Compte Jira"
            "Direction"      = "Direction"
            "Nom"            = "Nom"
            "Prénom"         = "Prénom"
            "Date Entrée"    = "Date Entrée"
            "Date PLD"       = "Date PLD"
            "Type ressource" = "Type ressource"
            "Statut"         = "Statut"
            "Tarif € TTC"   = "Tarif € TTC"
        }
        $dateAttrs = @("Date Entrée", "Date PLD")

        while (-not $isLast) {
            $url = '{0}?startAt={1}&maxResults={2}&includeAttributes=true' -f $assetsBaseUrl, $startAt, $maxResults
            try { $resp = Invoke-AssetsAqlPost -Url $url -Headers $jiraHeaders -AqlQuery $aql }
            catch { Write-ErrLog "Assets AQL échoué (startAt=$startAt)."; break }
            if (-not $resp) { break }

            # Dictionnaire id → nom depuis objectTypeAttributes (racine de la réponse)
            $attrDict = @{}
            foreach ($ota in @($resp.objectTypeAttributes)) {
                if ($ota.id -and $ota.name) {
                    $attrDict[[string]$ota.id] = Fix-DoubleUtf8 ([string]$ota.name)
                }
            }

            foreach ($v in @($resp.values)) {
                if (-not $v.id) { continue }
                $out = @{}

                foreach ($attr in @($v.attributes)) {
                    # Résolution du nom via dictionnaire (gateway v1)
                    $attrName = $null
                    if ($attr.objectTypeAttributeId) {
                        $attrName = $attrDict[[string]$attr.objectTypeAttributeId]
                    }
                    # Fallback ancien format (si objectTypeAttribute embarqué)
                    if (-not $attrName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                        $attrName = Fix-DoubleUtf8 ([string]$attr.objectTypeAttribute.name)
                    }
                    if (-not $attrName) { continue }

                    $outputKey = $null
                    foreach ($w in $wantedMap.Keys) {
                        if ($attrName -eq $w) { $outputKey = $wantedMap[$w]; break }
                    }
                    if (-not $outputKey) { continue }

                    $vals = $attr.objectAttributeValues

                    if ($outputKey -eq "Compte Jira") {
                        if ($vals -and $vals.Count -gt 0) {
                            $v0 = $vals[0]
                            if ($v0.displayValue) { $out["Compte Jira"] = Fix-DoubleUtf8 ([string]$v0.displayValue) }
                            if ($v0.user -and $v0.user.key) { $out["Compte Jira ID"] = [string]$v0.user.key }
                        }
                        continue
                    }

                    if ($vals -and $vals.Count -gt 0 -and $vals[0].displayValue) {
                        $value = Fix-DoubleUtf8 ([string]$vals[0].displayValue)
                        if ($dateAttrs -contains $outputKey) { $value = Normalize-AssetDate $value }
                        if ($outputKey -eq "Date PLD") { $out["Date Sortie"] = $value }
                        else { $out[$outputKey] = $value }
                    } else {
                        if ($outputKey -eq "Date PLD") { $out["Date Sortie"] = "" }
                    }
                }

                if (-not $out.ContainsKey("Date Sortie"))  { $out["Date Sortie"] = "" }
                if (-not $out.ContainsKey("Statut"))       { $out["Statut"] = "Inconnu" }
                if (-not $out.ContainsKey("Tarif € TTC"))  { $out["Tarif € TTC"] = "Non défini" }
                if (-not $out.ContainsKey("Prénom"))       { $out["Prénom"] = "" }
                if (-not $out.ContainsKey("Date Entrée"))  { $out["Date Entrée"] = "" }

                if ($out.Keys.Count -gt 0) { $assetUsers += [pscustomobject]$out }
            }

            $isLast  = [bool]$resp.isLast
            $startAt += $maxResults
            Write-Info "Assets: $($assetUsers.Count) users chargés (startAt=$startAt, isLast=$isLast)"
            if ($startAt -gt 500000) { Write-Warn "Assets pagination anormale."; break }
        }

        Save-Json -Path $assetsCacheFile -Object $assetUsers
        Write-Info "Assets users sauvegardés: $($assetUsers.Count)"
    }
}
Write-Info "Assets users: $($assetUsers.Count)"

# ======================================================================
# CHARGEMENT : ISSUES
# ======================================================================
$issuesById = @{}
if ($UseCache_Issues) {
    $cachedIssues = Load-Json -Path $issuesCacheFile
    if ($cachedIssues) {
        foreach ($p in $cachedIssues.PSObject.Properties) {
            $issuesById[[string]$p.Name] = $p.Value
        }
        Write-Info "Issues cache: $($issuesById.Count)"
    }
}

$issuesMissing = 0; $issuesSkipped = 0
foreach ($wl in $worklogs) {
    if (-not $wl.issue) { continue }
    $issueId = [string]$wl.issue.id
    $self    = [string]$wl.issue.self
    if ([string]::IsNullOrWhiteSpace($issueId) -or [string]::IsNullOrWhiteSpace($self)) { continue }
    if ($issuesById.ContainsKey($issueId)) { $issuesSkipped++; continue }
    $issuesMissing++
    try {
        $issue = Invoke-ApiGet -Url $self -Headers $jiraHeaders
        if ($issue -and $issue.id) { $issuesById[[string]$issue.id] = $issue }
    } catch { Write-Warn "Issue $self : $($_.Exception.Message)" }
}
Write-Info "Issues: $issuesSkipped en cache, $issuesMissing téléchargées, total=$($issuesById.Count)"

# Tempo account metadata cache
$tempoAccountMetaCache = @{}
if ($UseCache_TempoAccounts) {
    $c = Load-Json -Path $tempoAcctCache
    if ($c) { foreach ($p in $c.PSObject.Properties) { $tempoAccountMetaCache[$p.Name] = $p.Value } }
}

# ======================================================================
# CONSTRUCTION DES MAPS
# ======================================================================
$userTeamsMap    = Map-MembersToTeamsList -teamsData $teamsData
$userTeamExitMap = Map-MembersToTeamExitDates -teamsData $teamsData

$assetByAccountId = @{}
foreach ($u in $assetUsers) {
    $aid = [string]$u."Compte Jira ID"
    if (-not [string]::IsNullOrWhiteSpace($aid)) { $assetByAccountId[$aid] = $u }
}

$userWorklogs = @{}
foreach ($wl in $worklogs) {
    $aid = ""; try { $aid = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    if (-not $userWorklogs.ContainsKey($aid)) {
        $userWorklogs[$aid] = New-Object System.Collections.Generic.List[object]
    }
    $userWorklogs[$aid].Add($wl)
}

$timeLoggedMap = @{}
foreach ($wl in $worklogs) {
    $aid = ""; try { $aid = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    if (-not $timeLoggedMap.ContainsKey($aid)) { $timeLoggedMap[$aid] = 0.0 }
    $timeLoggedMap[$aid] += ([double]$wl.timeSpentSeconds) / 3600.0
}

$allUsers = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $userTeamsMap.Keys)  { [void]$allUsers.Add($k) }
foreach ($k in $userWorklogs.Keys)  { [void]$allUsers.Add($k) }

# ======================================================================
# APPEL DES 5 FEUILLES
# ======================================================================
Write-Info "--- Construction Feuille 1 ---"
$resultF1 = Build-WorklogsIssues `
    -AllUsers $allUsers `
    -UserTeamsMap $userTeamsMap `
    -UserWorklogs $userWorklogs `
    -IssuesById $issuesById `
    -AssetByAccountId $assetByAccountId `
    -TempoAccountMetaCache $tempoAccountMetaCache `
    -TempoHeaders $tempoHeaders `
    -JiraHeaders $jiraHeaders
Write-Info "Feuille 1: $($resultF1.Rows.Count) lignes"

# Sauvegarder caches mis à jour par F1
Save-Json -Path $tempoAcctCache -Object $tempoAccountMetaCache
$obj = [ordered]@{}
foreach ($k in $issuesById.Keys) { $obj[[string]$k] = $issuesById[[string]$k] }
Save-Json -Path $issuesCacheFile -Object $obj
Write-Info "Issues sauvegardées (avec parents): $($issuesById.Count)"

Write-Info "--- Construction Feuille 2 ---"
$resultF2 = Build-SaisiTempsTempo `
    -TeamsData $teamsData `
    -TimeLoggedMap $timeLoggedMap `
    -HeuresAttendue $heuresAttendue `
    -AssetByAccountId $assetByAccountId `
    -PeriodeFrom $From `
    -PeriodeTo $To `
    -JoursFeries $feries `
    -UserWorkloadDays $userWorkloadDays
Write-Info "Feuille 2: $($resultF2.Rows.Count) lignes"

Write-Info "--- Construction Feuille 3 ---"
$resultF3 = Build-SaisiTempsAsset `
    -Worklogs $worklogs `
    -UserTeamsMap $userTeamsMap `
    -AssetByAccountId $assetByAccountId `
    -WorkloadSchemeMap $workloadSchemeMap `
    -WeekBuckets $weekBuckets
Write-Info "Feuille 3: $($resultF3.Rows.Count) lignes"

Write-Info "--- Construction Feuille 4 ---"
$resultF4 = Build-Anomalies `
    -AssetUsers $assetUsers `
    -UserTeamsMap $userTeamsMap `
    -UserTeamExitMap $userTeamExitMap `
    -TimeLoggedMap $timeLoggedMap `
    -AssetByAccountId $assetByAccountId `
    -JiraBaseUrl $jiraBaseUrl `
    -JiraHeaders $jiraHeaders `
    -UseCacheGroups $UseCache_Groups `
    -GroupsCachePath $groupsCachePath `
    -UseCacheWorklogs $UseCache_Worklogs `
    -JiraUsersCachePath $jiraUsersCachePath `
    -PeriodeFrom $From `
    -PeriodeTo $To `
    -JoursFeries $feries `
    -UserWorkloadDays $userWorkloadDays
Write-Info "Feuille 4: $($resultF4.Rows.Count) lignes"

Write-Info "--- Construction Feuille 5 ---"
$resultF5 = Build-WorklogsSansDate `
    -Worklogs $worklogs `
    -IssuesById $issuesById `
    -UserTeamsMap $userTeamsMap `
    -JiraHeaders $jiraHeaders
Write-Info "Feuille 5: $($resultF5.Rows.Count) lignes"

# ======================================================================
# EXPORTS CSV
# ======================================================================
$csvF1 = Join-Path $csvSubDir "Worklogs & Issues.csv"
$csvF2 = Join-Path $csvSubDir "SaisiTempsTempo.csv"
$csvF3 = Join-Path $csvSubDir "SaisiTempsAsset.csv"
$csvF4 = Join-Path $csvSubDir "Sans équipe & sans saisie.csv"
$csvF5 = Join-Path $csvSubDir "Worklogs sans Start Date.csv"

Export-StrictCsv -Path $csvF1 -Headers $resultF1.Headers -Rows $resultF1.Rows
Export-StrictCsv -Path $csvF2 -Headers $resultF2.Headers -Rows $resultF2.Rows
Export-StrictCsv -Path $csvF3 -Headers $resultF3.Headers -Rows $resultF3.Rows
Export-StrictCsv -Path $csvF4 -Headers $resultF4.Headers -Rows $resultF4.Rows
Export-StrictCsv -Path $csvF5 -Headers $resultF5.Headers -Rows $resultF5.Rows

Write-Info "CSV générés dans: $csvSubDir"

# ======================================================================
# XLSX (v2.1 — ImportExcel sans COM, ou COM en STA avec timeout)
# ======================================================================
if ($CreateXlsx) {
    $xlsxName = "MVP-Reporting_{0}_{1}_{2}.xlsx" -f $fromStr, $toStr, $runStamp
    $xlsxPath = Join-Path $exportsDir $xlsxName
    $sheetToCsv = [ordered]@{
        "Worklogs & Issues"         = $csvF1
        "SaisiTempsTempo"           = $csvF2
        "SaisiTempsAsset"           = $csvF3
        "Sans équipe & sans saisie" = $csvF4
        "Worklogs sans Start Date"  = $csvF5
    }

    $xlsxGenerated = $false

    # ── Méthode 1 : Module ImportExcel (pas de COM, pas de blocage) ──────
    if (-not $xlsxGenerated) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
            Write-Info "XLSX: génération via ImportExcel..."
            # Supprimer le fichier s'il existe déjà
            if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
            foreach ($sheetName in $sheetToCsv.Keys) {
                $csvPath = $sheetToCsv[$sheetName]
                if (Test-Path $csvPath) {
                    $data = Import-Csv -Path $csvPath -Delimiter ";" -Encoding UTF8
                    if ($data.Count -gt 0) {
                        $data | Export-Excel -Path $xlsxPath -WorksheetName $sheetName -AutoSize -Append
                    } else {
                        # Feuille vide avec juste les headers
                        $emptyRow = [pscustomobject]@{ Info = "(aucune donnée)" }
                        $emptyRow | Export-Excel -Path $xlsxPath -WorksheetName $sheetName -Append
                    }
                }
            }
            $xlsxGenerated = $true
            Write-Info "XLSX OK (ImportExcel): $xlsxPath"
        } catch {
            Write-Warn "ImportExcel non disponible ou erreur: $($_.Exception.Message)"
            Write-Info "Tentative via COM Excel en STA..."
        }
    }

    # ── Méthode 2 : COM Excel en STA runspace avec timeout 120s ──────────
    if (-not $xlsxGenerated) {
        Write-Info "XLSX: génération via COM Excel (STA, timeout 120s)..."
        $staScript = {
            param($SheetToCsv, $XlsxPath)
            $excel = $null
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $wb = $excel.Workbooks.Add()
                # Supprimer les feuilles par défaut sauf la première
                while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
                $firstSheet = $true
                foreach ($sheetName in $SheetToCsv.Keys) {
                    $csvPath = $SheetToCsv[$sheetName]
                    if (-not (Test-Path $csvPath)) { continue }
                    if ($firstSheet) {
                        $ws = $wb.Sheets.Item(1)
                        $ws.Name = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))
                        $firstSheet = $false
                    } else {
                        $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
                        $ws.Name = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))
                    }
                    # Lire CSV et peupler
                    $lines = [System.IO.File]::ReadAllLines($csvPath, [System.Text.Encoding]::UTF8)
                    for ($r = 0; $r -lt $lines.Count; $r++) {
                        $cols = $lines[$r].Split(";")
                        for ($c = 0; $c -lt $cols.Count; $c++) {
                            $ws.Cells.Item($r + 1, $c + 1).Value2 = $cols[$c]
                        }
                    }
                    # Autofit
                    [void]$ws.UsedRange.Columns.AutoFit()
                }
                $wb.SaveAs($XlsxPath, 51) # 51 = xlOpenXMLWorkbook
                $wb.Close($false)
                return "OK"
            } catch {
                return "ERREUR: $($_.Exception.Message)"
            } finally {
                if ($excel) {
                    try { $excel.Quit() } catch {}
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
                }
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }
        }

        # Exécuter dans un runspace STA avec timeout
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $runspace.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()

        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript($staScript)
        [void]$ps.AddArgument($sheetToCsv)
        [void]$ps.AddArgument($xlsxPath)

        $handle = $ps.BeginInvoke()
        $timeoutSec = 120
        $completed = $handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($timeoutSec))

        if ($completed) {
            $result = $ps.EndInvoke($handle)
            if ($result -and $result[0] -eq "OK") {
                $xlsxGenerated = $true
                Write-Info "XLSX OK (COM STA): $xlsxPath"
            } else {
                Write-Warn "XLSX COM erreur: $($result -join ' ')"
            }
        } else {
            Write-Warn "XLSX COM timeout après ${timeoutSec}s — abandon"
            $ps.Stop()
            # Tuer Excel orphelin si lancé
            Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
                Where-Object { $_.StartTime -gt (Get-Date).AddSeconds(-$timeoutSec - 10) } |
                Stop-Process -Force -ErrorAction SilentlyContinue
        }
        $ps.Dispose()
        $runspace.Close()
        $runspace.Dispose()
    }

    # ── Méthode 3 : Ancien appel (fonction Common.ps1) ───────────────────
    if (-not $xlsxGenerated) {
        Write-Warn "XLSX: toutes les méthodes modernes ont échoué."
        Write-Warn "Essai de la méthode legacy New-ExcelWorkbookFromCsvSheets..."
        try {
            New-ExcelWorkbookFromCsvSheets -SheetToCsvPath $sheetToCsv -XlsxPath $xlsxPath
            $xlsxGenerated = $true
            Write-Info "XLSX OK (legacy): $xlsxPath"
        } catch {
            Write-Warn "XLSX: ÉCHEC COMPLET — $($_.Exception.Message)"
            Write-Warn "Les CSV sont disponibles dans: $csvSubDir"
        }
    }
}

# ======================================================================
# RÉSUMÉ
# ======================================================================
Write-Host ""
Write-Info "============ RÉSUMÉ v2.1 ============"
Write-Info "Période                           : $($From.ToString('dd/MM/yyyy')) -> $($To.ToString('dd/MM/yyyy'))"
Write-Info "Jours ouvrés (hors fériés)        : $joursOuvres j"
Write-Info "Heures attendues base (8h/j)      : $heuresAttendue h"
Write-Info "Workload schemes personnalisés    : $($userWorkloadDays.Count) users"
Write-Host ""
for ($wi = 0; $wi -lt 5; $wi++) {
    $b = $weekBuckets[$wi]
    if ($null -ne $b.Start) {
        Write-Info "Semaine $($wi+1)                        : $($b.Start.ToString('dd/MM')) -> $($b.End.ToString('dd/MM'))"
    } else {
        Write-Info "Semaine $($wi+1)                        : (vide)"
    }
}
Write-Host ""
Write-Info "Worklogs                          : $($worklogs.Count)"
Write-Info "Issues                            : $($issuesById.Count)"
Write-Info "Teams                             : $($teamsData.Count)"
Write-Info "Workload schemes (noms)           : $($workloadSchemeMap.Count) users"
Write-Info "Workload schemes (jours)          : $($userWorkloadDays.Count) users"
Write-Info "Assets users                      : $($assetUsers.Count)"
Write-Host ""
Write-Info "Feuille 1 - Worklogs & Issues     : $($resultF1.Rows.Count) lignes"
Write-Info "Feuille 2 - SaisiTempsTempo       : $($resultF2.Rows.Count) lignes"
Write-Info "Feuille 3 - SaisiTempsAsset       : $($resultF3.Rows.Count) lignes"
Write-Info "Feuille 4 - Anomalies             : $($resultF4.Rows.Count) lignes"
Write-Info "Feuille 5 - Worklogs sans date    : $($resultF5.Rows.Count) lignes"
Write-Host ""
Write-Info "Cache rafraîchi                   : $(if($refreshList.Count -gt 0){$refreshList -join ', '}else{'(aucun)'})"
Write-Info "Cache conservé                    : $(if($cachedList.Count -gt 0){$cachedList -join ', '}else{'(aucun)'})"
Write-Host ""
Write-Info "CSV:  $csvSubDir"
if ($CreateXlsx -and $xlsxGenerated) { Write-Info "XLSX: $xlsxPath" }
elseif ($CreateXlsx) { Write-Warn "XLSX: non généré (voir warnings ci-dessus)" }
Write-Info "Log:  $logFile"

Write-Log "=== FIN EXÉCUTION MVP-REPORTING v2.1 ===" "INFO"
