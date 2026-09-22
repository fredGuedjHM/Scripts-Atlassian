<#
.SYNOPSIS
    Audit complet et catégorisation d'arbitrage des applications connectées Jira Cloud.
.DESCRIPTION
    - MODE STRICTEMENT DRY-RUN (LECTURE SEULE).
    - Catégorisation automatique :
        * Socle Système Atlassian (Obligatoire)
        * Connecteur Optionnel Atlassian (Slack, Teams, Opsgenie, etc.)
        * Application Métier Tierce (Tempo, Xray, Figma, Links Hierarchy, etc.)
    - Détection de l'empreinte de données et évaluation de l'impact en cas de suppression.
    - Export CSV UTF-8 avec BOM dans exports/.
#>

[CmdletBinding()]
param(
    [string]$ProxyUrl = "http://prc37cti1.hm.dm.ad:8080"
)

# ======================================================================
# 1. INITIALISATION PROXY, TLS & ENCODAGE
# ======================================================================
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if (-not $ProxyUrl) {
    try {
        $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $testUri = New-Object System.Uri("https://api.atlassian.com")
        $pUri = $sysProxy.GetProxy($testUri)
        if ($pUri -and $pUri.AbsoluteUri -ne $testUri.AbsoluteUri) { $ProxyUrl = $pUri.AbsoluteUri }
    } catch {}
}

function Repair-TextEncoding([string]$str) {
    if ([string]::IsNullOrWhiteSpace($str)) { return "" }
    try {
        if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
            $bytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        }
    } catch {}
    return $str
}

function Invoke-SafeGet {
    param([string]$Url, [hashtable]$Headers, [int]$TimeoutSec = 8)
    $params = @{ Uri = $Url; Method = "Get"; Headers = $Headers; TimeoutSec = $TimeoutSec }
    if ($ProxyUrl) {
        $params["Proxy"] = $ProxyUrl
        $params["ProxyUseDefaultCredentials"] = $true
    }
    try {
        return (Invoke-RestMethod @params)
    } catch {
        return $null
    }
}

# ======================================================================
# 2. DOSSIERS & IDENTIFIANTS
# ======================================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$exportsDir = Join-Path $scriptDir "exports"
if (-not (Test-Path $exportsDir)) { [void](New-Item -ItemType Directory -Path $exportsDir -Force) }

function Write-Info($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [WARN] $msg" -ForegroundColor Yellow }

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable: $jiraCredFile" }

$jiraData    = Import-Clixml -Path $jiraCredFile
$baseUrl     = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

$pair        = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${jiraEmail}:${jiraToken}"))
$authHeaders = @{
    "Authorization" = "Basic $pair"
    "Accept"        = "application/json"
}

$upmHeaders = @{
    "Authorization" = "Basic $pair"
    "Accept"        = "application/vnd.atl.plugins.installed+json"
}

Write-Host ""
Write-Host "======================================================================" -ForegroundColor Magenta
Write-Host " [DRY-RUN] AUDIT & CATÉGORISATION DES APPLICATIONS CONNECTÉES" -ForegroundColor Magenta
Write-Host " Instance : $baseUrl" -ForegroundColor Magenta
Write-Host "======================================================================" -ForegroundColor Magenta
Write-Host ""

$appsDict = @{}

# Bruit technique pur à ignorer
$pureInternalNoise = '(?i)^(com\.atlassian\.jcs|jquery|\$\{project|com\.pyxis\.greenhopper|jira\.polaris|read-only-string)'

# ======================================================================
# ÉTAPE 1 : SCAN UPM (PLUGINS MARKETPLACE & CONNECT)
# ======================================================================
Write-Info "1/3 Scan des modules UPM..."

$upmResp = Invoke-SafeGet -Url "$baseUrl/rest/plugins/1.0/?include-built-in=false" -Headers $upmHeaders
if ($upmResp -and $upmResp.plugins) {
    foreach ($p in $upmResp.plugins) {
        $k = [string]$p.key
        $n = Repair-TextEncoding ([string]$p.name)
        $v = if ($p.vendor -and $p.vendor.name) { Repair-TextEncoding ([string]$p.vendor.name) } else { "Atlassian" }

        if ($k -match $pureInternalNoise -or $n -match '^\$\{') { continue }

        $isUser = if ($p.userInstalled) { "Oui (Administrateur)" } else { "Souscription Cloud / Système" }

        $appsDict[$k] = [pscustomobject]@{
            Nom                  = $n
            Cle                  = $k
            Editeur              = $v
            CategorieModule      = ""
            InstalleParUser      = $isUser
            Empreinte            = "Module actif"
            Preconisation        = ""
            ImpactDesinstallation= ""
        }
    }
}

# ======================================================================
# ÉTAPE 2 : SCAN CONNECT ADDONS (SaaS / OAuth)
# ======================================================================
Write-Info "2/3 Scan des intégrations Connect & SaaS..."

$connectResp = Invoke-SafeGet -Url "$baseUrl/rest/atlassian-connect/1/addons" -Headers $authHeaders
if ($connectResp -and $connectResp.addons) {
    foreach ($a in $connectResp.addons) {
        $k = [string]$a.key
        $n = Repair-TextEncoding ([string]$a.name)
        $v = if ($a.vendor -and $a.vendor.name) { Repair-TextEncoding ([string]$a.vendor.name) } else { "Connect SaaS" }

        if ($k -match $pureInternalNoise -or $n -match '^\$\{') { continue }

        if (-not $appsDict.ContainsKey($k)) {
            $appsDict[$k] = [pscustomobject]@{
                Nom                  = $n
                Cle                  = $k
                Editeur              = $v
                CategorieModule      = ""
                InstalleParUser      = "Oui (OAuth / SaaS)"
                Empreinte            = "Intégration externe"
                Preconisation        = ""
                ImpactDesinstallation= ""
            }
        }
    }
}

# ======================================================================
# ÉTAPE 3 : DÉTECTION XRAY & TEMPO VIA MODÈLE DE DONNÉES
# ======================================================================
Write-Info "3/3 Détection approfondie des applications métiers clés (Xray, Tempo)..."

# A. Détection Xray (via types de tickets natifs Xray)
$issueTypes = @(Invoke-SafeGet -Url "$baseUrl/rest/api/3/issuetype" -Headers $authHeaders)
$xrayTypesFound = @()
foreach ($it in $issueTypes) {
    $desc = [string]$it.description
    $name = [string]$it.name
    if ($desc -match '(?i)Xray' -or $name -match '^(Test|Test Execution|Test Plan|Test Set|Precondition|Sub Test Execution)$') {
        $xrayTypesFound += $name
    }
}

if ($xrayTypesFound.Count -gt 0) {
    $appsDict["com.xpandit.plugins.xray"] = [pscustomobject]@{
        Nom                  = "Xray Test Management"
        Cle                  = "com.xpandit.plugins.xray"
        Editeur              = "Idevio / Xpand IT"
        CategorieModule      = "Application Métier Tierce"
        InstalleParUser      = "Souscription Cloud (Organisation)"
        Empreinte            = "$($xrayTypesFound.Count) types de tickets Xray ($($xrayTypesFound -join ', '))"
        Preconisation        = "🟢 CONSERVER (Outil Stratégique - Tests QA)"
        ImpactDesinstallation= "ÉLEVÉ : Perte de la traçabilité des plans et exécutions de tests QA."
    }
}

# B. Détection Tempo (via champs Forge & Custom Fields)
$fields = @(Invoke-SafeGet -Url "$baseUrl/rest/api/3/field" -Headers $authHeaders)
$tempoFields = @()
foreach ($f in $fields) {
    if ($f.custom -and $f.schema -and $f.schema.custom) {
        $cType = [string]$f.schema.custom
        $fName = Repair-TextEncoding ([string]$f.name)
        if ($cType -match 'fa75e928-007a-4af4-9530-76503bcd4cba' -or $cType -match '(?i)tempo' -or $fName -match '(?i)^Tempo|Account') {
            $tempoFields += $fName
        }
    }
}

if ($tempoFields.Count -gt 0) {
    $appsDict["io.tempo.jira"] = [pscustomobject]@{
        Nom                  = "Tempo Timesheets & Planner"
        Cle                  = "io.tempo.jira"
        Editeur              = "Tempo Software"
        CategorieModule      = "Application Métier Tierce"
        InstalleParUser      = "Souscription Cloud (Organisation)"
        Empreinte            = "$($tempoFields.Count) champs actifs ($($tempoFields -join ', '))"
        Preconisation        = "🟢 CONSERVER (Outil Stratégique - Saisie des Temps & Budgets)"
        ImpactDesinstallation= "CRITIQUE : Arrêt de la saisie des temps et du suivi budgétaire des projets."
    }
}

# ======================================================================
# ÉTAPE 4 : CALCUL AUTOMATIQUE DES CATÉGORIES & PRÉCONISATIONS
# ======================================================================
foreach ($k in $appsDict.Keys) {
    $app = $appsDict[$k]
    if ($app.Preconisation) { continue }

    $kLower = $app.Cle.ToLower()
    $nLower = $app.Nom.ToLower()
    $vLower = $app.Editeur.ToLower()

    # 1. MODULES DU SOCLE SYSTÈME ATLASSIAN (OBLIGATOIRE DE GARDER)
    if ($kLower -match '(?i)(streams|toolkit|proforma|servicedesk|roadmaps|jpo|teams|inline-create|workflow-designer|atlaskit|less)') {
        $app.CategorieModule       = "🛑 Socle Système Atlassian"
        $app.Preconisation         = "🟢 CONSERVER OBLIGATOIREMENT (Composant socle Jira Cloud)"
        $app.ImpactDesinstallation = "CRITIQUE : Dysfonctionnement de l'interface ou des fonctionnalités natives Jira."
    }
    # 2. CONNECTEURS OPTIONNELS ATLASSIAN (À ARBITRER SELON USAGES)
    elseif ($vLower -match '(?i)atlassian' -or $kLower -match '(?i)(slack|teams|opsgenie|statuspage)') {
        $app.CategorieModule = "🔵 Connecteur Optionnel Atlassian"

        if ($kLower.Contains("slack")) {
            $app.Preconisation         = "⚪ À DÉSINSTALLER si l'entreprise utilise Teams et non Slack"
            $app.ImpactDesinstallation = "FAIBLE : Suppression des notifications vers les canaux Slack."
        }
        elseif ($kLower.Contains("teams") -or $nLower.Contains("teams")) {
            $app.Preconisation         = "🔵 CONSERVER si notifications / création de tickets via Teams"
            $app.ImpactDesinstallation = "MOYEN : Fin de l'intégration des tickets dans les canaux Microsoft Teams."
        }
        elseif ($kLower.Contains("opsgenie")) {
            $app.Preconisation         = "🔵 CONSERVER si gestion d'astreintes / alertes Opsgenie"
            $app.ImpactDesinstallation = "MOYEN : Perte de la synchronisation des alertes d'incidents Opsgenie."
        }
        elseif ($kLower.Contains("statuspage")) {
            $app.Preconisation         = "🔵 CONSERVER si communication des incidents via Statuspage"
            $app.ImpactDesinstallation = "FAIBLE : Fin de la liaison directe tickets Jira <-> incidents Statuspage."
        }
        else {
            $app.Preconisation         = "🟡 À ARBITRER (Vérifier si le service connecté est utilisé)"
            $app.ImpactDesinstallation = "FAIBLE : Perte de l'interconnexion avec le service externe."
        }
    }
    # 3. APPLICATIONS MÉTIERS TIERCES (MARKETPLACE / FORGE)
    else {
        $app.CategorieModule = "🟢 Application Métier Tierce"

        if ($kLower.Contains("figma") -or $nLower.Contains("figma")) {
            $app.Preconisation         = "🔵 CONSERVER si les équipes UI/UX lient des maquettes dans Jira"
            $app.ImpactDesinstallation = "FAIBLE : Les aperçus interactifs de maquettes Figma ne s'afficheront plus."
        }
        elseif ($kLower.Contains("links") -or $nLower.Contains("links hierarchy")) {
            $app.Preconisation         = "🟢 CONSERVER si l'arborescence visuelle de tickets est exploitée"
            $app.ImpactDesinstallation = "MOYEN : Perte du panneau de visualisation hiérarchique des liens."
        }
        elseif ($kLower.Contains("drawio") -or $kLower.Contains("gliffy")) {
            $app.Preconisation         = "🔵 CONSERVER si création de schémas / diagrammes dans les tickets"
            $app.ImpactDesinstallation = "MOYEN : Impossibilité de modifier les diagrammes existants."
        }
        else {
            $app.Preconisation         = "⚪ CANDIDATE DÉSINSTALLATION (Vérifier absence d'usage métier)"
            $app.ImpactDesinstallation = "À ÉVALUER : Vérifier avec les utilisateurs avant suppression de licence."
        }
    }
}

# ======================================================================
# 5. EXPORT DU RAPPORT
# ======================================================================
$finalList = @($appsDict.Values | Sort-Object CategorieModule, Nom)

$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportCsv = Join-Path $exportsDir "Audit_Connected_Apps_Categorise_${dateStamp}.csv"

$headers = @(
    "Nom", "Cle", "Editeur", "CategorieModule", "InstalleParUser",
    "Empreinte", "Preconisation", "ImpactDesinstallation"
)

$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers -join ";"))

foreach ($r in $finalList) {
    $lineVals = @()
    foreach ($h in $headers) {
        $val = [string]$r.$h
        if ($val -match '[;"\r\n]') { $val = '"' + ($val -replace '"', '""') + '"' }
        $lineVals += $val
    }
    $csvLines.Add(($lineVals -join ";"))
}

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($exportCsv, $csvLines, $utf8WithBom)

Write-Host ""
Write-Info "=== AUDIT TERMINÉ AVEC SUCCÈS ==="
Write-Info "Fichier exporté : $exportCsv"
Write-Host ""

$finalList | Select-Object Nom, Editeur, CategorieModule, Preconisation | Format-Table -AutoSize