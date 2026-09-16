<#
audit-users-assets.ps1
- Compare les utilisateurs Jira actifs avec les fiches Jira Assets
- Croise avec Tempo pour afficher le cumul des heures saisies sur la période
- Filtre: uniquement @harmonie-mutuelle.fr et @prestataire.sihm.fr
- Classifie par Direction (DSIM / DSIT / Gouvernance / Autre) selon les groupes Jira
- 1 ligne par user, groupes concaténés
- Calcul jours ouvrés hors fériés français, base 7h/jour
- Génère 2 CSV:
    CSV 1 - Audit anomalies: users sans fiche Asset ou fiche inactive avec saisie
    CSV 2 - Contrôle saisie: users avec fiche Asset active, écart vs heures attendues

Règles de filtrage CSV 1:
  - Fiche Asset active             → exclu (va dans CSV 2)
  - Fiche Asset inactive + 0h      → exclu (vraiment parti)
  - Fiche Asset inactive + >0h     → INCLUS (anomalie)
  - Pas de fiche Asset             → INCLUS

Credentials: réutilise les mêmes fichiers que mvp-reporting.ps1
    .\secrets\jira-jiradot.cred.xml
    .\secrets\tempo-token.xml

Lancement:
    .\audit-users-assets.ps1                    # popup mois, cache activé
    .\audit-users-assets.ps1 -UseCache:$false   # rechargement complet
    .\audit-users-assets.ps1 -From 2026-04-01 -To 2026-04-30  # période imposée
#>

[CmdletBinding()]
param(
    [datetime]$From,
    [datetime]$To,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [bool]$UseCache = $true,
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl
)

# ----------------------------
# Dossiers
# ----------------------------
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$cacheDir   = Join-Path $scriptDir "cache"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $secretsDir; Ensure-Dir $cacheDir; Ensure-Dir $logsDir; Ensure-Dir $exportsDir

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ----------------------------
# Log
# ----------------------------
$logFile = Join-Path $logsDir ("audit-users-assets_{0}.log" -f $runStamp)

function Write-Log {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $logFile -Value "[$ts] [$Level] $Message" -ErrorAction Stop } catch {}
}
function Write-Info($msg) {
    Write-Host "[INFO] $msg"
    if ($msg) { Write-Log $msg "INFO" }
}
function Write-Warn($msg)   { Write-Warning $msg;        Write-Log $msg "WARN" }
function Write-ErrLog($msg) { Write-Error $msg;          Write-Log $msg "ERROR" }

Write-Log "=== DÉBUT EXÉCUTION AUDIT-USERS-ASSETS ===" "INFO"

# ----------------------------
# Proxy
# ----------------------------
function Initialize-Proxy {
    param([switch]$UseSystemProxy, [string]$ProxyUrl)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($ProxyUrl) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl,$true)
        } elseif ($UseSystemProxy) {
            [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        } else {
            [System.Net.WebRequest]::DefaultWebProxy = $null
        }
    } catch { Write-Warn "Init proxy: $($_.Exception.Message)" }
}
Initialize-Proxy -UseSystemProxy:$UseSystemProxy -ProxyUrl $ProxyUrl

function Get-EffectiveProxyUri {
    param([Parameter(Mandatory)][string]$TargetUrl)
    if ($ProxyUrl -and $ProxyUrl.Trim()) { return $ProxyUrl }
    if (-not $UseSystemProxy) { return $null }
    try { $dest = [uri]$TargetUrl } catch { return $null }
    $wp = [System.Net.WebRequest]::DefaultWebProxy
    if (-not $wp -or $wp.IsBypassed($dest)) { return $null }
    $proxy = $wp.GetProxy($dest)
    if (-not $proxy -or $proxy.AbsoluteUri -eq $dest.AbsoluteUri) { return $null }
    return $proxy.AbsoluteUri
}

# ----------------------------
# WebException body
# ----------------------------
function Get-WebExceptionBody([System.Net.WebException]$ex) {
    try {
        if (-not $ex.Response) { return $null }
        $s = $ex.Response.GetResponseStream(); $r = New-Object System.IO.StreamReader($s)
        $b = $r.ReadToEnd(); $r.Dispose(); $s.Dispose(); return $b
    } catch { return $null }
}

# ----------------------------
# Jours fériés français
# ----------------------------
function Get-JoursFeries([int]$Annee) {
    # Jours fériés fixes
    $fixes = @(
        (Get-Date -Year $Annee -Month 1 -Day 1),   # Nouvel An
        (Get-Date -Year $Annee -Month 5 -Day 1),   # Fête du Travail
        (Get-Date -Year $Annee -Month 5 -Day 8),   # Victoire 1945
        (Get-Date -Year $Annee -Month 7 -Day 14),  # Fête nationale
        (Get-Date -Year $Annee -Month 8 -Day 15),  # Assomption
        (Get-Date -Year $Annee -Month 11 -Day 1),  # Toussaint
        (Get-Date -Year $Annee -Month 11 -Day 11), # Armistice
        (Get-Date -Year $Annee -Month 12 -Day 25)  # Noël
    )
    # Pâques (algorithme de Meeus)
    $a = $Annee % 19
    $b = [math]::Floor($Annee / 100)
    $c = $Annee % 100
    $dd = [math]::Floor($b / 4)
    $e = $b % 4
    $f = [math]::Floor(($b + 8) / 25)
    $g = [math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $dd - $g + 15) % 30
    $i = [math]::Floor($c / 4)
    $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [math]::Floor(($a + 11 * $h + 22 * $l) / 451)
    $mois = [math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $jour = (($h + $l - 7 * $m + 114) % 31) + 1
    $paques = Get-Date -Year $Annee -Month $mois -Day $jour

    $lundiPaques    = $paques.AddDays(1)   # Lundi de Pâques
    $ascension      = $paques.AddDays(39)  # Ascension
    $lundiPentecote = $paques.AddDays(50)  # Lundi de Pentecôte

    $all = $fixes + @($lundiPaques, $ascension, $lundiPentecote)
    return ($all | ForEach-Object { $_.Date })
}

# ----------------------------
# Sélection du mois via popup (Windows Forms)
# ----------------------------
function Show-MonthPickerDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Audit Users/Assets - Sélection du mois"
    $form.Size = New-Object System.Drawing.Size(350, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Sélectionnez le mois à auditer :"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($label)

    $comboMois = New-Object System.Windows.Forms.ComboBox
    $comboMois.DropDownStyle = "DropDownList"
    $comboMois.Location = New-Object System.Drawing.Point(20, 50)
    $comboMois.Size = New-Object System.Drawing.Size(140, 25)
    $moisNoms = @("Janvier","Février","Mars","Avril","Mai","Juin",
                   "Juillet","Août","Septembre","Octobre","Novembre","Décembre")
    foreach ($m in $moisNoms) { $comboMois.Items.Add($m) | Out-Null }
    $today = Get-Date
    $moisPrecedent = $today.AddMonths(-1).Month
    $comboMois.SelectedIndex = $moisPrecedent - 1
    $form.Controls.Add($comboMois)

    $comboAnnee = New-Object System.Windows.Forms.ComboBox
    $comboAnnee.DropDownStyle = "DropDownList"
    $comboAnnee.Location = New-Object System.Drawing.Point(180, 50)
    $comboAnnee.Size = New-Object System.Drawing.Size(80, 25)
    $anneeActuelle = $today.Year
    foreach ($a in @(($anneeActuelle - 1), $anneeActuelle, ($anneeActuelle + 1))) {
        $comboAnnee.Items.Add($a) | Out-Null
    }
    $anneeMoisPrec = $today.AddMonths(-1).Year
    $comboAnnee.SelectedItem = $anneeMoisPrec
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
        Write-Info "Audit annulé par l'utilisateur."; exit 0
    }

    $moisChoisi   = $comboMois.SelectedIndex + 1
    $anneeChoisie = [int]$comboAnnee.SelectedItem
    $form.Dispose()

    $debut = Get-Date -Year $anneeChoisie -Month $moisChoisi -Day 1
    $fin   = $debut.AddMonths(1).AddDays(-1)
    return [pscustomobject]@{ From = $debut.Date; To = $fin.Date }
}

# ----------------------------
# Période
# ----------------------------
if (-not $From -or -not $To) {
    $periode = Show-MonthPickerDialog
    $From = $periode.From; $To = $periode.To
}
if ($From -gt $To) { throw "From > To ($From > $To)" }
$fromStr = $From.ToString("yyyy-MM-dd"); $toStr = $To.ToString("yyyy-MM-dd")
Write-Info "Période: $($From.ToString('dd/MM/yyyy')) -> $($To.ToString('dd/MM/yyyy'))"

# ----------------------------
# Jira creds
# ----------------------------
$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) {
    throw "Fichier Jira creds introuvable: $jiraCredFile`nLance d'abord Save-JiraCredential.ps1."
}
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

function New-BasicAuthHeader([string]$User,[string]$Pass) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$User`:$Pass"))
    return @{ Authorization="Basic $b64"; Accept="application/json" }
}
$jiraHeaders = New-BasicAuthHeader -User $jiraEmail -Pass $jiraToken
Write-Info "Jira: $jiraBaseUrl (user=$jiraEmail)"

# ----------------------------
# Tempo token
# ----------------------------
$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"

function ConvertFrom-SecureStringToPlain([securestring]$Secure) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}
function Get-TempoToken {
    if (Test-Path $tempoTokenFile) {
        $obj = Import-Clixml -Path $tempoTokenFile
        if ($obj -and $obj.Token) { return (ConvertFrom-SecureStringToPlain -Secure $obj.Token) }
    }
    $sec = Read-Host -AsSecureString -Prompt "Tempo API token (masqué)"
    if (-not $sec) { throw "Tempo token vide." }
    [pscustomobject]@{ Token=$sec } | Export-Clixml -Path $tempoTokenFile
    return (ConvertFrom-SecureStringToPlain -Secure $sec)
}
$tempoToken = Get-TempoToken
$tempoHeaders = @{ Authorization="Bearer $tempoToken"; Accept="application/json" }

# ----------------------------
# HTTP wrappers
# ----------------------------
function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][hashtable]$Headers)
    Write-Info "GET $Url"
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        if ($body) { Write-ErrLog "GET $Url : $($_.Exception.Message)`nBody:`n$body" }
        else       { Write-ErrLog "GET $Url : $($_.Exception.Message)" }
        throw
    }
}

function Invoke-AssetsAqlPost {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][hashtable]$Headers, [Parameter(Mandatory)][string]$AqlQuery)
    Write-Info "POST $Url"
    $jsonBody = '{"qlQuery": "' + $AqlQuery + '"}'
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
    $params = @{ Method='POST'; Uri=$Url; Headers=$Headers; Body=$bytes; ContentType='application/json; charset=utf-8'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        if ($body) { Write-ErrLog "POST Assets $Url : $($_.Exception.Message)`nBody:`n$body" }
        else       { Write-ErrLog "POST Assets $Url : $($_.Exception.Message)" }
        throw
    }
}

# ----------------------------
# Fix double encodage UTF-8
# ----------------------------
function Fix-DoubleUtf8([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return $s }
    try {
        $latin1 = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $bytes  = $latin1.GetBytes($s)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    } catch { return $s }
}

# ----------------------------
# Cache helpers
# ----------------------------
$usersCacheFile      = Join-Path $cacheDir "audit_jira_users.json"
$userGroupsCacheFile = Join-Path $cacheDir "audit_user_groups.json"
$assetsCacheFile     = Join-Path $cacheDir "saisi_temps_asset.json"
$worklogCacheFile    = Join-Path $cacheDir "worklogs.json"

function Save-Json([string]$Path, $Object) {
    $Object | ConvertTo-Json -Depth 80 | Set-Content -Path $Path -Encoding UTF8
}
function Load-Json([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

# ----------------------------
# Tempo worklogs
# ----------------------------
function Get-TempoWorklogs([string]$FromStr,[string]$ToStr) {
    $all = @(); $url = "https://api.eu.tempo.io/4/worklogs?from=$FromStr&to=$ToStr"
    while ($true) {
        $resp = Invoke-ApiGet -Url $url -Headers $tempoHeaders
        if ($resp.results) { $all += $resp.results }
        if ($resp.metadata -and $resp.metadata.next) { $url = $resp.metadata.next } else { break }
    }
    return $all
}

# ----------------------------
# Jira users (filtrés)
# ----------------------------
function Get-AllActiveJiraUsers {
    $users = @(); $startAt = 0; $maxResults = 200
    while ($true) {
        $url = "$jiraBaseUrl/rest/api/3/users/search?startAt=$startAt&maxResults=$maxResults"
        $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders
        if (-not $resp -or $resp.Count -eq 0) { break }
        foreach ($u in $resp) {
            if ($u.accountType -eq "atlassian" -and $u.active -eq $true) {
                $email = if ($u.emailAddress) { [string]$u.emailAddress } else { "" }
                if ($email -notmatch "@harmonie-mutuelle\.fr$" -and $email -notmatch "@prestataire\.sihm\.fr$") { continue }
                $users += [pscustomobject]@{
                    AccountId   = [string]$u.accountId
                    DisplayName = [string]$u.displayName
                    Email       = $email
                    Active      = [bool]$u.active
                }
            }
        }
        if ($resp.Count -lt $maxResults) { break }
        $startAt += $maxResults
    }
    return $users
}

# ----------------------------
# Jira user groups
# ----------------------------
function Get-UserJiraGroups([string]$AccountId) {
    $url = "$jiraBaseUrl/rest/api/3/user/groups?accountId=$AccountId"
    try {
        $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders
        $groupNames = @()
        foreach ($g in $resp) { if ($g.name) { $groupNames += [string]$g.name } }
        return $groupNames
    } catch {
        Write-Warn "Groupes pour $AccountId : $($_.Exception.Message)"
        return @()
    }
}

# ----------------------------
# Direction depuis les groupes
# ----------------------------
function Get-DirectionFromGroups([string[]]$Groups) {
    $directions = @()
    foreach ($g in $Groups) {
        if ($g -match "(?i)dsim")        { if ($directions -notcontains "DSIM")        { $directions += "DSIM" } }
        if ($g -match "(?i)dsit")        { if ($directions -notcontains "DSIT")        { $directions += "DSIT" } }
        if ($g -match "(?i)gouvernance") { if ($directions -notcontains "Gouvernance") { $directions += "Gouvernance" } }
    }
    if ($directions.Count -eq 0) { return "Autre" }
    return ($directions -join " / ")
}

# ----------------------------
# Assets
# ----------------------------
function Get-AssetsObjectIds {
    $ids = New-Object System.Collections.Generic.List[string]
    $startAt = 0; $maxResults = 50; $isLast = $false
    $base = "https://api.atlassian.com/jsm/assets/workspace/$AssetsWorkspaceId/v1/object/aql"
    $aql = '(objectType = Employé OR objectType = Prestataire)'
    while (-not $isLast) {
        $url = '{0}?startAt={1}&maxResults={2}&includeAttributes=true' -f $base, $startAt, $maxResults
        try { $resp = Invoke-AssetsAqlPost -Url $url -Headers $jiraHeaders -AqlQuery $aql }
        catch { Write-ErrLog "Assets AQL échoué (startAt=$startAt)."; break }
        if (-not $resp) { break }
        foreach ($v in @($resp.values)) { if ($v.id) { $ids.Add([string]$v.id) } }
        $isLast = [bool]$resp.isLast; $startAt += $maxResults
        if ($startAt -gt 500000) { break }
    }
    return $ids
}

function Get-AssetsUserDataByObjectId {
    param([Parameter(Mandatory)][string]$ObjectId)
    $url = "https://api.atlassian.com/jsm/assets/workspace/$AssetsWorkspaceId/v1/object/$ObjectId"
    $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders

    $out = @{}
    $wantedMap = @{
        "Compte Jira"="Compte Jira"; "Direction"="Direction"; "Nom"="Nom"; "Prénom"="Prénom"
        "Type ressource"="Type ressource"; "Statut"="Statut"
    }

    foreach ($attr in ($resp.attributes | ForEach-Object { $_ })) {
        $fixedName = Fix-DoubleUtf8 $attr.objectTypeAttribute.name
        $vals = $attr.objectAttributeValues

        $outputKey = $null
        foreach ($wanted in $wantedMap.Keys) {
            if ($fixedName -eq $wanted) { $outputKey = $wantedMap[$wanted]; break }
        }
        if (-not $outputKey) { continue }

        if ($outputKey -eq "Compte Jira") {
            if ($vals -and $vals.Count -gt 0) {
                $v0 = $vals[0]
                if ($v0.displayValue) { $out["Compte Jira"] = Fix-DoubleUtf8 ([string]$v0.displayValue) }
                if ($v0.user -and $v0.user.key) { $out["Compte Jira ID"] = [string]$v0.user.key }
            }
            continue
        }

        if ($vals -and $vals.Count -gt 0 -and $vals[0].displayValue) {
            $out[$outputKey] = Fix-DoubleUtf8 ([string]$vals[0].displayValue)
        }
    }

    if (-not $out.ContainsKey("Statut"))        { $out["Statut"] = "" }
    if (-not $out.ContainsKey("Prénom"))         { $out["Prénom"] = "" }
    if (-not $out.ContainsKey("Nom"))            { $out["Nom"] = "" }
    if (-not $out.ContainsKey("Direction"))      { $out["Direction"] = "" }
    if (-not $out.ContainsKey("Type ressource")) { $out["Type ressource"] = "" }

    return $out
}

function Get-AllAssetsUsers {
    if ($UseCache) {
        $cached = Load-Json -Path $assetsCacheFile
        if ($cached) { Write-Info "Assets cache: $($cached.Count)"; return $cached }
    }
    $ids = Get-AssetsObjectIds
    Write-Info "Assets IDs: $($ids.Count)"
    $list = @(); $count = 0
    foreach ($id in $ids) {
        try {
            $u = Get-AssetsUserDataByObjectId -ObjectId $id
            if ($u -and $u.Keys.Count -gt 0) { $list += [pscustomobject]$u }
            $count++
            if ($count % 100 -eq 0) { Write-Info "Assets: $count / $($ids.Count)" }
        } catch { Write-Warn "Assets object $id : $($_.Exception.Message)" }
    }
    Save-Json -Path $assetsCacheFile -Object $list
    Write-Info "Assets sauvegardés: $($list.Count)"
    return $list
}
# ======================================================================
# EXÉCUTION PRINCIPALE
# ======================================================================

# --- Étape 1 : Utilisateurs Jira actifs ---
$jiraUsers = $null
if ($UseCache) { $jiraUsers = Load-Json -Path $usersCacheFile }
if (-not $jiraUsers) {
    $jiraUsers = Get-AllActiveJiraUsers
    Save-Json -Path $usersCacheFile -Object $jiraUsers
    Write-Info "Utilisateurs Jira actifs (filtrés): $($jiraUsers.Count)"
} else { Write-Info "Utilisateurs Jira actifs (cache): $($jiraUsers.Count)" }

# --- Étape 2a : Assets ---
$assetUsers = Get-AllAssetsUsers
Write-Info "Fiches Assets: $($assetUsers.Count)"

$assetByAccountId = @{}
foreach ($u in $assetUsers) {
    $aid = [string]$u."Compte Jira ID"
    if (-not [string]::IsNullOrWhiteSpace($aid)) { $assetByAccountId[$aid] = $u }
}

# --- Étape 2b : Worklogs Tempo ---
$worklogs = $null
if ($UseCache) { $worklogs = Load-Json -Path $worklogCacheFile }
if (-not $worklogs) {
    $worklogs = Get-TempoWorklogs -FromStr $fromStr -ToStr $toStr
    Save-Json -Path $worklogCacheFile -Object $worklogs
    Write-Info "Worklogs récupérés: $($worklogs.Count)"
} else { Write-Info "Worklogs cache: $($worklogs.Count)" }

$timeLoggedMap = @{}
foreach ($wl in $worklogs) {
    $aid = ""; try { $aid = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    if (-not $timeLoggedMap.ContainsKey($aid)) { $timeLoggedMap[$aid] = 0.0 }
    $timeLoggedMap[$aid] += ([double]$wl.timeSpentSeconds) / 3600.0
}
Write-Info "Utilisateurs avec temps saisi: $($timeLoggedMap.Count)"

# --- Étape 2c : Jours ouvrés hors fériés et heures attendues ---
$feries = Get-JoursFeries -Annee $From.Year
if ($To.Year -ne $From.Year) { $feries += Get-JoursFeries -Annee $To.Year }

$joursOuvres = 0
$d = $From
while ($d -le $To) {
    if ($d.DayOfWeek -ne [DayOfWeek]::Saturday -and $d.DayOfWeek -ne [DayOfWeek]::Sunday -and $feries -notcontains $d.Date) {
        $joursOuvres++
    }
    $d = $d.AddDays(1)
}
$heuresAttendue = $joursOuvres * 8  # 8h/jour standard
Write-Info "Jours ouvrés (hors fériés): $joursOuvres ($heuresAttendue h attendues à 8h/j)"

# --- Étape 3 : Séparer les users en 2 listes ---

# CSV 1 : anomalies (pas de fiche OU fiche inactive avec saisie)
$problematicUsers = @()

# CSV 2 : contrôle saisie (fiche active)
$activeAssetUsers = @()

foreach ($u in $jiraUsers) {
    $aid = [string]$u.AccountId
    $hasAsset = $assetByAccountId.ContainsKey($aid)
    $hours = 0.0; if ($timeLoggedMap.ContainsKey($aid)) { $hours = [double]$timeLoggedMap[$aid] }

    if ($hasAsset) {
        $assetStatut = [string]$assetByAccountId[$aid].Statut
        $asset = $assetByAccountId[$aid]

        if ($assetStatut -match "(?i)^actif$") {
            # Fiche active → CSV 2
            $activeAssetUsers += [pscustomobject]@{
                AccountId     = $aid
                DisplayName   = [string]$u.DisplayName
                Email         = [string]$u.Email
                Active        = [bool]$u.Active
                HeuresSaisies = [math]::Round($hours, 2)
                Nom           = [string]$asset.Nom
                Prenom        = [string]$asset."Prénom"
                Direction     = [string]$asset.Direction
                TypeRessource = [string]$asset."Type ressource"
            }
            continue
        }

        # Fiche inactive + 0h → exclu (vraiment parti)
        if ($hours -eq 0) {
            Write-Log "Exclu (inactif sans saisie): $($u.DisplayName) [$aid]" "DEBUG"
            continue
        }

        # Fiche inactive + >0h → anomalie → CSV 1
        $statutLabel = if ([string]::IsNullOrWhiteSpace($assetStatut)) { "Statut vide" } else { $assetStatut }
    } else {
        # Pas de fiche → CSV 1
        $statutLabel = "Pas de fiche Asset"
    }

    $problematicUsers += [pscustomobject]@{
        AccountId     = $aid
        DisplayName   = [string]$u.DisplayName
        Email         = [string]$u.Email
        StatutAsset   = $statutLabel
        HasAsset      = $hasAsset
        Active        = [bool]$u.Active
        HeuresSaisies = [math]::Round($hours, 2)
    }
}
Write-Info "CSV 1 - Anomalies: $($problematicUsers.Count)"
Write-Info "CSV 2 - Actifs à contrôler: $($activeAssetUsers.Count)"

# --- Étape 4 : Récupérer les groupes de TOUS les users ---
$userGroupsMap = @{}
if ($UseCache) {
    $cachedGroups = Load-Json -Path $userGroupsCacheFile
    if ($cachedGroups) {
        foreach ($p in $cachedGroups.PSObject.Properties) { $userGroupsMap[$p.Name] = @($p.Value) }
        Write-Info "Groupes utilisateurs (cache): $($userGroupsMap.Count)"
    }
}

$allUsersToFetchGroups = New-Object System.Collections.Generic.HashSet[string]
foreach ($u in $problematicUsers)  { [void]$allUsersToFetchGroups.Add($u.AccountId) }
foreach ($u in $activeAssetUsers)  { [void]$allUsersToFetchGroups.Add($u.AccountId) }

$groupsFetched = 0
foreach ($aid in $allUsersToFetchGroups) {
    if ($userGroupsMap.ContainsKey($aid)) { continue }
    $groups = Get-UserJiraGroups -AccountId $aid
    $userGroupsMap[$aid] = $groups
    $groupsFetched++
    if ($groupsFetched % 50 -eq 0) { Write-Info "Groupes récupérés: $groupsFetched" }
}
if ($groupsFetched -gt 0) {
    Save-Json -Path $userGroupsCacheFile -Object $userGroupsMap
    Write-Info "Groupes sauvegardés: $($userGroupsMap.Count) utilisateurs"
}

# --- Étape 5 : CSV 1 — Audit anomalies (1 ligne par user) ---
$csvRows = New-Object System.Collections.Generic.List[object]

foreach ($u in $problematicUsers) {
    $aid = $u.AccountId
    $groups = @()
    if ($userGroupsMap.ContainsKey($aid)) { $groups = @($userGroupsMap[$aid]) }
    if ($groups.Count -eq 0) { $groups = @("Aucun groupe") }

    $direction = Get-DirectionFromGroups -Groups $groups
    $groupesConcat = ($groups | Sort-Object) -join " | "

    $nomAsset = ""; $prenomAsset = ""; $directionAsset = ""; $typeRes = ""
    if ($u.HasAsset -and $assetByAccountId.ContainsKey($aid)) {
        $asset = $assetByAccountId[$aid]
        $nomAsset       = [string]$asset.Nom
        $prenomAsset    = [string]$asset."Prénom"
        $directionAsset = [string]$asset.Direction
        $typeRes        = [string]$asset."Type ressource"
    }

    $assetStatutBrut = if ($u.HasAsset -and $assetByAccountId.ContainsKey($aid)) { [string]$assetByAccountId[$aid].Statut } else { "" }

    $csvRows.Add([pscustomobject]@{
        "Direction"         = $direction
        "Groupes Jira"      = $groupesConcat
        "Account ID"        = $aid
        "Nom Complet"       = $u.DisplayName
        "Email"             = $u.Email
        "Statut Jira"       = if ($u.Active) { "Actif" } else { "Inactif" }
        "Situation"         = if ($u.HasAsset) { "Fiche existante" } else { "Pas de fiche Asset" }
        "Statut Asset"      = if ($u.HasAsset) { $assetStatutBrut } else { "" }
        "Temps saisi (h)"   = $u.HeuresSaisies
        "Nom (Asset)"       = $nomAsset
        "Prénom (Asset)"    = $prenomAsset
        "Direction (Asset)" = $directionAsset
        "Type Ressource"    = $typeRes
    }) | Out-Null
}

$csvRowsSorted = $csvRows | Sort-Object -Property @(
    @{Expression="Direction"; Ascending=$true},
    @{Expression="Groupes Jira"; Ascending=$true},
    @{Expression="Nom Complet"; Ascending=$true}
)

# --- Étape 5b : CSV 2 — Contrôle saisie actifs (1 ligne par user) ---
$csvRowsActif = New-Object System.Collections.Generic.List[object]

foreach ($u in $activeAssetUsers) {
    $aid = $u.AccountId
    $groups = @()
    if ($userGroupsMap.ContainsKey($aid)) { $groups = @($userGroupsMap[$aid]) }
    if ($groups.Count -eq 0) { $groups = @("Aucun groupe") }

    $direction = Get-DirectionFromGroups -Groups $groups
    $groupesConcat = ($groups | Sort-Object) -join " | "

    $ecart = [math]::Round($u.HeuresSaisies - $heuresAttendue, 2)
    $statutSaisie = if ($u.HeuresSaisies -eq 0) { "Aucune saisie" }
                    elseif ($u.HeuresSaisies -lt $heuresAttendue) { "Incomplet" }
                    elseif ($u.HeuresSaisies -eq $heuresAttendue) { "Complet" }
                    else { "Dépassement" }

    $csvRowsActif.Add([pscustomobject]@{
        "Direction"         = $direction
        "Groupes Jira"      = $groupesConcat
        "Account ID"        = $aid
        "Nom Complet"       = $u.DisplayName
        "Email"             = $u.Email
        "Nom (Asset)"       = $u.Nom
        "Prénom (Asset)"    = $u.Prenom
        "Direction (Asset)" = $u.Direction
        "Type Ressource"    = $u.TypeRessource
        "Temps saisi (h)"   = $u.HeuresSaisies
        "Temps attendu (h)" = $heuresAttendue
        "Écart (h)"         = $ecart
        "Statut Saisie"     = $statutSaisie
    }) | Out-Null
}

$csvRowsActifSorted = $csvRowsActif | Sort-Object -Property @(
    @{Expression="Statut Saisie"; Ascending=$true},
    @{Expression="Direction"; Ascending=$true},
    @{Expression="Groupes Jira"; Ascending=$true},
    @{Expression="Nom Complet"; Ascending=$true}
)

# --- Étape 6 : Exports CSV ---
$csvHeaders = @("Direction","Groupes Jira","Account ID","Nom Complet","Email",
                "Statut Jira","Situation","Statut Asset","Temps saisi (h)",
                "Nom (Asset)","Prénom (Asset)","Direction (Asset)","Type Ressource")

$csvHeadersActif = @("Direction","Groupes Jira","Account ID","Nom Complet","Email",
                      "Nom (Asset)","Prénom (Asset)","Direction (Asset)","Type Ressource",
                      "Temps saisi (h)","Temps attendu (h)","Écart (h)","Statut Saisie")

$csvPath = Join-Path $exportsDir ("Audit-Users-Assets_{0}_{1}_{2}.csv" -f $fromStr, $toStr, $runStamp)
$csvPathActif = Join-Path $exportsDir ("Controle-Saisie-Actifs_{0}_{1}_{2}.csv" -f $fromStr, $toStr, $runStamp)

$orderedRows = foreach ($r in $csvRowsSorted) {
    $o = [ordered]@{}; foreach ($h in $csvHeaders) { $o[$h] = $r.$h }; [pscustomobject]$o
}
$orderedRows | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $csvPath
Write-Info "CSV 1 généré: $csvPath"

$orderedRowsActif = foreach ($r in $csvRowsActifSorted) {
    $o = [ordered]@{}; foreach ($h in $csvHeadersActif) { $o[$h] = $r.$h }; [pscustomobject]$o
}
$orderedRowsActif | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $csvPathActif
Write-Info "CSV 2 généré: $csvPathActif"

# --- Résumé ---
$countDSIM  = ($csvRows | Where-Object { $_.Direction -match "DSIM" }).Count
$countDSIT  = ($csvRows | Where-Object { $_.Direction -match "DSIT" }).Count
$countGouv  = ($csvRows | Where-Object { $_.Direction -match "Gouvernance" }).Count
$countAutre = ($csvRows | Where-Object { $_.Direction -eq "Autre" }).Count
$countSansFiche = ($problematicUsers | Where-Object { $_.StatutAsset -eq "Pas de fiche Asset" }).Count
$countInactif   = ($problematicUsers | Where-Object { $_.StatutAsset -ne "Pas de fiche Asset" }).Count

$countAucune      = ($activeAssetUsers | Where-Object { $_.HeuresSaisies -eq 0 }).Count
$countIncomplet   = ($activeAssetUsers | Where-Object { $_.HeuresSaisies -gt 0 -and $_.HeuresSaisies -lt $heuresAttendue }).Count
$countComplet     = ($activeAssetUsers | Where-Object { $_.HeuresSaisies -eq $heuresAttendue }).Count
$countDepassement = ($activeAssetUsers | Where-Object { $_.HeuresSaisies -gt $heuresAttendue }).Count

Write-Host ""
Write-Info "============ RÉSUMÉ ============"
Write-Info "Période                           : $($From.ToString('dd/MM/yyyy')) -> $($To.ToString('dd/MM/yyyy'))"
Write-Info "Jours ouvrés (hors fériés)        : $joursOuvres j"
Write-Info "Heures attendues (7h/j)           : $heuresAttendue h"
Write-Host ""
Write-Info "--- Utilisateurs Jira (filtrés) ---"
Write-Info "Total actifs                      : $($jiraUsers.Count)"
Write-Info "Avec fiche Asset active           : $($activeAssetUsers.Count)"
Write-Info "Sans fiche Asset                  : $countSansFiche"
Write-Info "Fiche Asset non active + saisie   : $countInactif"
Write-Host ""
Write-Info "--- CSV 1 : Audit anomalies ---"
Write-Info "Utilisateurs                      : $($problematicUsers.Count)"
Write-Info "  dont DSIM                       : $countDSIM"
Write-Info "  dont DSIT                       : $countDSIT"
Write-Info "  dont Gouvernance                : $countGouv"
Write-Info "  dont Autre                      : $countAutre"
Write-Host ""
Write-Info "--- CSV 2 : Contrôle saisie actifs ---"
Write-Info "Utilisateurs                      : $($activeAssetUsers.Count)"
Write-Info "  Aucune saisie                   : $countAucune"
Write-Info "  Incomplet (< $heuresAttendue h)          : $countIncomplet"
Write-Info "  Complet ($heuresAttendue h)               : $countComplet"
Write-Info "  Dépassement (> $heuresAttendue h)         : $countDepassement"
Write-Host ""
Write-Info "CSV 1 : $csvPath"
Write-Info "CSV 2 : $csvPathActif"
Write-Info "Log   : $logFile"

Write-Log "=== FIN EXÉCUTION AUDIT-USERS-ASSETS ===" "INFO"
