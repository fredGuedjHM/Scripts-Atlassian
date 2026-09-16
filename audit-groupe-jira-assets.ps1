<#
.SYNOPSIS
    Audit des habilitations et données RH pour les membres d'un groupe Jira donné (G1).
.DESCRIPTION
    - Gestion robuste du proxy d'entreprise et des caractères accentués UTF-8 (BOM inclus).
    - Normalisation stricte des dates d'entrée et de sortie au format DD/MM/YYYY.
    - Inspection dynamique des attributs du schéma Assets (Référentiel Personne).
    - Demande le nom du groupe Jira G1 cible.
    - Récupère pour chaque membre :
        * Autres groupes Jira (séparés par " | ")
        * Équipe(s) Tempo
        * Données Assets : Direction, Nom, Prénom, Date d'arrivée, Date de sortie, Motif de sortie, Affectation 2, Manager
    - Exporte le résultat en CSV (séparateur ;) UTF-8 avec BOM.
#>

[CmdletBinding()]
param(
    [string]$GroupName,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string]$ProxyUrl = "http://prc37cti1.hm.dm.ad:8080"
)

# ======================================================================
# 1. INITIALISATION PROXY, TLS & ENCODAGE CONSOLE
# ======================================================================
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $ProxyUrl) {
    try {
        $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $testUri = New-Object System.Uri("https://api.atlassian.com")
        $pUri = $sysProxy.GetProxy($testUri)
        if ($pUri -and $pUri.AbsoluteUri -ne $testUri.AbsoluteUri) {
            $ProxyUrl = $pUri.AbsoluteUri
        }
    } catch {}
}

# Fonction anti double-encodage UTF-8
function Fix-DoubleUtf8([string]$str) {
    if ([string]::IsNullOrWhiteSpace($str)) { return "" }
    try {
        if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
            $bytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        }
    } catch {}
    return $str
}

# Fonction de formatage des dates au format strict DD/MM/YYYY
function Format-DateFr([string]$rawDate) {
    if ([string]::IsNullOrWhiteSpace($rawDate)) { return "" }
    $rawDate = $rawDate.Trim()
    
    if ($rawDate -match '^\d{2}/\d{2}/\d{4}$') { return $rawDate }
    
    $dt = [datetime]::MinValue
    $formats = @(
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-ddTHH:mm:ss.fffZ",
        "yyyy-MM-ddTHH:mm:ssZ",
        "yyyy/MM/dd",
        "d/M/yyyy",
        "dd/MM/yyyy",
        "dd-MM-yyyy",
        "yyyy-MM-dd HH:mm:ss"
    )
    
    foreach ($fmt in $formats) {
        if ([datetime]::TryParseExact($rawDate, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString("dd/MM/yyyy")
        }
    }

    if ([datetime]::TryParse($rawDate, [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"), [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString("dd/MM/yyyy")
    }
    
    if ([datetime]::TryParse($rawDate, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString("dd/MM/yyyy")
    }
    
    return $rawDate
}

# Wrapper d'appels API
function Invoke-SafeGet {
    param([string]$Url, [hashtable]$Headers)
    $params = @{
        Uri     = $Url
        Method  = "Get"
        Headers = $Headers
    }
    if ($ProxyUrl) {
        $params["Proxy"] = $ProxyUrl
        $params["ProxyUseDefaultCredentials"] = $true
    }
    return (Invoke-RestMethod @params)
}

function Invoke-SafePostJson {
    param([string]$Url, [hashtable]$Headers, [string]$BodyJson)
    $params = @{
        Uri         = $Url
        Method      = "Post"
        Headers     = $Headers
        Body        = $BodyJson
        ContentType = "application/json; charset=utf-8"
    }
    if ($ProxyUrl) {
        $params["Proxy"] = $ProxyUrl
        $params["ProxyUseDefaultCredentials"] = $true
    }
    return (Invoke-RestMethod @params)
}

# ======================================================================
# 2. DOSSIERS & LOGS
# ======================================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$exportsDir = Join-Path $scriptDir "exports"

if (-not (Test-Path $exportsDir)) { [void](New-Item -ItemType Directory -Path $exportsDir -Force) }

function Write-Info($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [WARN] $msg" -ForegroundColor Yellow }
function Write-ErrLog($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [ERROR] $msg" -ForegroundColor Red }

# ======================================================================
# 3. CREDENTIALS JIRA & TEMPO
# ======================================================================
$jiraCredFile   = Join-Path $secretsDir "jira-jiradot.cred.xml"
$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"

if (-not (Test-Path $jiraCredFile)) {
    throw "Fichier Jira creds introuvable: $jiraCredFile"
}

$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

$pair        = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${jiraEmail}:${jiraToken}"))
$jiraHeaders = @{
    "Authorization" = "Basic $pair"
    "Accept"        = "application/json"
}

$tempoHeaders = $null
if (Test-Path $tempoTokenFile) {
    $tObj = Import-Clixml -Path $tempoTokenFile
    if ($tObj -and $tObj.Token) {
        $sec = $tObj.Token
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        $tPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        $tempoHeaders = @{
            "Authorization" = "Bearer $tPlain"
            "Accept"        = "application/json"
        }
    }
}

# ======================================================================
# 4. SÉLECTION DU GROUPE JIRA (G1)
# ======================================================================
if ([string]::IsNullOrWhiteSpace($GroupName)) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Audit Groupe Jira & Référentiel Personne"
    $form.Size = New-Object System.Drawing.Size(460, 180)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Entrez le nom exact du groupe Jira (G1) :"
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(400, 20)
    [void]$form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(20, 45)
    $txt.Size = New-Object System.Drawing.Size(400, 25)
    [void]$form.Controls.Add($txt)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Valider"
    $btnOK.Location = New-Object System.Drawing.Point(140, 85)
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    [void]$form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(230, 85)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    [void]$form.Controls.Add($btnCancel)

    $txt.Focus()
    $res = $form.ShowDialog()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($txt.Text)) {
        $GroupName = $txt.Text.Trim()
    } else {
        $GroupName = Read-Host "Nom du groupe Jira cible"
    }
    $form.Dispose()
}

if ([string]::IsNullOrWhiteSpace($GroupName)) {
    Write-ErrLog "Aucun groupe spécifié. Fin du script."
    exit 1
}

$GroupName = Fix-DoubleUtf8 $GroupName
Write-Info "Groupe Jira cible : $GroupName"

# ======================================================================
# 5. INSPECTION DES ATTRIBUTS ASSETS (CMDB RP)
# ======================================================================
Write-Info "Assets: inspection et cartographie de tous les attributs du Référentiel Personne..."

$assetsBaseUrl = "$jiraBaseUrl/gateway/api/jsm/assets/workspace/$AssetsWorkspaceId/v1/object/aql"
$aqlQuery      = '(objectType = Employé OR objectType = Prestataire)'
$initUrl       = "$assetsBaseUrl`?startAt=0&maxResults=50&includeAttributes=true"
$bodyJson      = @{ qlQuery = $aqlQuery } | ConvertTo-Json -Compress

$attrIdToName = @{}
try {
    $metaResp = Invoke-SafePostJson -Url $initUrl -Headers $jiraHeaders -BodyJson $bodyJson
    if ($metaResp.objectTypeAttributes) {
        foreach ($ota in $metaResp.objectTypeAttributes) {
            $attrIdToName[[string]$ota.id] = Fix-DoubleUtf8 ([string]$ota.name)
        }
    }
    Write-Info "Assets: $($attrIdToName.Count) attributs détectés dans le schéma RP."
} catch {
    Write-Warn "Impossible d'inspecter les métadonnées : $($_.Exception.Message)"
}

# ======================================================================
# 6. CHARGEMENT DU RÉFÉRENTIEL PERSONNE (ASSETS)
# ======================================================================
Write-Info "Assets: extraction des fiches RP..."
$assetUsersByAccountId = @{}
$assetUsersByName      = @{}

$startAt = 0
$maxRes  = 50
$isLast  = $false

while (-not $isLast) {
    $url = "$assetsBaseUrl`?startAt=$startAt&maxResults=$maxRes&includeAttributes=true"
    try {
        $resp = Invoke-SafePostJson -Url $url -Headers $jiraHeaders -BodyJson $bodyJson
    } catch {
        Write-ErrLog "Erreur Assets à startAt=$startAt : $($_.Exception.Message)"
        break
    }

    if (-not $resp -or -not $resp.values) { break }

    foreach ($obj in $resp.values) {
        $props = @{
            "Nom"             = ""
            "Prénom"          = ""
            "Direction"       = ""
            "Date d'arrivée"  = ""
            "Date de sortie"  = ""
            "Motif de sortie" = ""
            "Affectation 2"   = ""
            "Manager"         = ""
            "Compte Jira ID"  = ""
            "Compte Jira"     = ""
        }

        foreach ($att in $obj.attributes) {
            $aName = $null
            if ($att.objectTypeAttributeId -and $attrIdToName.ContainsKey([string]$att.objectTypeAttributeId)) {
                $aName = $attrIdToName[[string]$att.objectTypeAttributeId]
            } elseif ($att.objectTypeAttribute -and $att.objectTypeAttribute.name) {
                $aName = Fix-DoubleUtf8 ([string]$att.objectTypeAttribute.name)
            }
            if (-not $aName) { continue }

            $vals = $att.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) { continue }
            $valStr = Fix-DoubleUtf8 ([string]$vals[0].displayValue)

            switch -Wildcard ($aName) {
                "*Direction*"     { $props["Direction"] = $valStr }
                "Nom"             { $props["Nom"] = $valStr }
                "Prénom"          { $props["Prénom"] = $valStr }
                "*Entrée*"        { $props["Date d'arrivée"] = Format-DateFr $valStr }
                "*Arrivée*"       { $props["Date d'arrivée"] = Format-DateFr $valStr }
                "*PLD*"           { $props["Date de sortie"]  = Format-DateFr $valStr }
                "*Sortie*"        { if (-not $props["Date de sortie"]) { $props["Date de sortie"] = Format-DateFr $valStr } }
                "*Motif*"         { $props["Motif de sortie"] = $valStr }
                "*Affectation 2*" { $props["Affectation 2"] = $valStr }
                "*Manager*"       { $props["Manager"] = $valStr }
                "*Compte Jira*" {
                    $props["Compte Jira"] = $valStr
                    if ($vals[0].searchValue) { $props["Compte Jira ID"] = [string]$vals[0].searchValue }
                    elseif ($vals[0].user -and $vals[0].user.key) { $props["Compte Jira ID"] = [string]$vals[0].user.key }
                }
            }
        }

        $rowObj = [pscustomobject]$props

        if ($rowObj."Compte Jira ID") {
            $assetUsersByAccountId[[string]$rowObj."Compte Jira ID"] = $rowObj
        }
        $keyName = "$($rowObj.Prénom) $($rowObj.Nom)".Trim().ToLower()
        if ($keyName) { $assetUsersByName[$keyName] = $rowObj }
    }

    $isLast = if ($null -ne $resp.isLast) { [bool]$resp.isLast } else { $resp.values.Count -lt $maxRes }
    $startAt += $maxRes
}
Write-Info "Assets: $($assetUsersByAccountId.Count) fiches RP indexées par AccountId."

# ======================================================================
# 7. CHARGEMENT DES ÉQUIPES TEMPO
# ======================================================================
$userTempoTeamsMap = @{}
if ($tempoHeaders) {
    Write-Info "Tempo: extraction des équipes et de leurs membres..."
    try {
        $tUrl = "https://api.tempo.io/4/teams"
        while ($tUrl) {
            $tResp = Invoke-SafeGet -Url $tUrl -Headers $tempoHeaders
            foreach ($team in $tResp.results) {
                $teamName = Fix-DoubleUtf8 ([string]$team.name)
                $mUrl = "$($team.self)/members"
                try {
                    $mResp = Invoke-SafeGet -Url $mUrl -Headers $tempoHeaders
                    foreach ($m in $mResp.results) {
                        if ($m.member -and $m.member.accountId) {
                            $mAccId = [string]$m.member.accountId
                            if (-not $userTempoTeamsMap.ContainsKey($mAccId)) {
                                $userTempoTeamsMap[$mAccId] = New-Object System.Collections.Generic.List[string]
                            }
                            if (-not $userTempoTeamsMap[$mAccId].Contains($teamName)) {
                                $userTempoTeamsMap[$mAccId].Add($teamName)
                            }
                        }
                    }
                } catch {}
            }
            $tUrl = if ($tResp.metadata -and $tResp.metadata.next) { [string]$tResp.metadata.next } else { $null }
        }
        Write-Info "Tempo: cartographie des équipes terminée."
    } catch {
        Write-Warn "Impossible d'extraire les équipes Tempo : $($_.Exception.Message)"
    }
}

# ======================================================================
# 8. EXTRACTION DES MEMBRES DU GROUPE G1 & RÉCUPÉRATION DES GROUPES
# ======================================================================
Write-Info "Jira: recherche des membres du groupe '$GroupName'..."
$members = New-Object System.Collections.ArrayList
$startAt = 0
$maxRes  = 50

while ($true) {
    $encGroup = [Uri]::EscapeDataString($GroupName)
    $gUrl = "$jiraBaseUrl/rest/api/3/group/member?groupname=$encGroup&startAt=$startAt&maxResults=$maxRes&includeInactiveUsers=true"
    try {
        $gResp = Invoke-SafeGet -Url $gUrl -Headers $jiraHeaders
    } catch {
        Write-ErrLog "Erreur lors de la récupération des membres du groupe '$GroupName' : $($_.Exception.Message)"
        break
    }

    if (-not $gResp -or -not $gResp.values -or $gResp.values.Count -eq 0) { break }

    foreach ($u in $gResp.values) {
        [void]$members.Add($u)
    }

    if ($gResp.isLast -or $gResp.values.Count -lt $maxRes) { break }
    $startAt += $gResp.values.Count
}

Write-Info "Jira: $($members.Count) membre(s) trouvé(s) dans le groupe '$GroupName'."

if ($members.Count -eq 0) {
    Write-Warn "Aucun utilisateur à traiter. Fin."
    exit 0
}

# ======================================================================
# 9. CROISEMENT ET COMPILATION DU RAPPORT
# ======================================================================
Write-Info "Compilation du rapport et analyse des autres groupes..."
$reportRows = New-Object System.Collections.ArrayList
$count = 0

foreach ($user in $members) {
    $count++
    $accId       = [string]$user.accountId
    $displayName = Fix-DoubleUtf8 ([string]$user.displayName)
    $email       = [string]$user.emailAddress
    $isActive    = [bool]$user.active

    # 1. Autres groupes Jira
    $otherGroups = @()
    try {
        $uUrl = "$jiraBaseUrl/rest/api/3/user?accountId=$accId&expand=groups"
        $uDetail = Invoke-SafeGet -Url $uUrl -Headers $jiraHeaders
        if ($uDetail.groups -and $uDetail.groups.items) {
            $otherGroups = @($uDetail.groups.items | ForEach-Object { Fix-DoubleUtf8 ([string]$_.name) } | Where-Object { $_ -ne $GroupName })
        }
    } catch {
        try {
            $ugUrl = "$jiraBaseUrl/rest/api/3/user/groups?accountId=$accId"
            $ugDetail = Invoke-SafeGet -Url $ugUrl -Headers $jiraHeaders
            $otherGroups = @($ugDetail | ForEach-Object { Fix-DoubleUtf8 ([string]$_.name) } | Where-Object { $_ -ne $GroupName })
        } catch {}
    }

    $otherGroupsStr = if ($otherGroups.Count -gt 0) { ($otherGroups | Sort-Object) -join " | " } else { "(aucun autre groupe)" }

    # 2. Équipe Tempo
    $tempoTeamStr = if ($userTempoTeamsMap.ContainsKey($accId)) {
        ($userTempoTeamsMap[$accId] | Sort-Object) -join " | "
    } else {
        "(aucune équipe Tempo)"
    }

    # 3. Réconciliation RP Assets
    $rp = $null
    if ($assetUsersByAccountId.ContainsKey($accId)) {
        $rp = $assetUsersByAccountId[$accId]
    } elseif ($displayName -and $assetUsersByName.ContainsKey($displayName.Trim().ToLower())) {
        $rp = $assetUsersByName[$displayName.Trim().ToLower()]
    }

    $row = [pscustomobject]@{
        "Groupe Habilitation (G1)" = $GroupName
        "AccountId"                = $accId
        "DisplayName"              = $displayName
        "Email"                    = $email
        "Compte Actif"             = if ($isActive) { "Oui" } else { "Non" }
        "Autres Groupes Jira"      = $otherGroupsStr
        "Équipe Tempo"             = $tempoTeamStr
        "Direction"                = if ($rp) { $rp.Direction } else { "" }
        "Nom (RP)"                 = if ($rp) { $rp.Nom } else { "" }
        "Prénom (RP)"              = if ($rp) { $rp.Prénom } else { "" }
        "Date d'arrivée"           = if ($rp) { Format-DateFr $rp."Date d'arrivée" } else { "" }
        "Date de sortie"           = if ($rp) { Format-DateFr $rp."Date de sortie" } else { "" }
        "Motif de sortie"          = if ($rp) { $rp."Motif de sortie" } else { "" }
        "Affectation 2"            = if ($rp) { $rp."Affectation 2" } else { "" }
        "Manager"                  = if ($rp) { $rp.Manager } else { "" }
    }

    [void]$reportRows.Add($row)
    Write-Info "[$count/$($members.Count)] Traité : $displayName"
}

# ======================================================================
# 10. EXPORT CSV AVEC BOM UTF-8 STRICT (POUR EXCEL)
# ======================================================================
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeGroupName = $GroupName -replace '[\\/:*?"<>| ]', '_'
$exportCsv = Join-Path $exportsDir "Audit_Groupe_${safeGroupName}_${dateStamp}.csv"

$headers = @(
    "Groupe Habilitation (G1)", "AccountId", "DisplayName", "Email", "Compte Actif",
    "Autres Groupes Jira", "Équipe Tempo", "Direction", "Nom (RP)", "Prénom (RP)",
    "Date d'arrivée", "Date de sortie", "Motif de sortie", "Affectation 2", "Manager"
)

$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers -join ";"))

foreach ($r in $reportRows) {
    $lineVals = @()
    foreach ($h in $headers) {
        $val = [string]$r.$h
        if ($val -match '[;"\r\n]') {
            $val = '"' + ($val -replace '"', '""') + '"'
        }
        $lineVals += $val
    }
    $csvLines.Add(($lineVals -join ";"))
}

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($exportCsv, $csvLines, $utf8WithBom)

Write-Host ""
Write-Info "=== AUDIT TERMINÉ AVEC SUCCÈS ==="
Write-Info "Fichier généré : $exportCsv"
Write-Host ""

$reportRows | Select-Object "DisplayName", "Direction", "Date d'arrivée", "Date de sortie", "Manager" | Format-Table -AutoSize