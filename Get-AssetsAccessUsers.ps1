# ============================================================
# Nom          : Get-AssetsAccessUsers.ps1
# Description  : Identification des utilisateurs ayant un
#                accès applicatif (licence JSM / droits produit)
#                ou des rôles sur les schémas d'objets Assets
#                dans l'instance Jiradot.
#                Les comptes applicatifs (accountType=app) sont
#                exclus du resultat et traces en avertissement.
# Version      : 1.2
# Date         : 11/09/2026
# Auteur       : DSIM / HM_DSIM_PACT
# Prerequis    : Fichier secrets/jira-jiradot.cred.xml valide
# Usage        : .\Get-AssetsAccessUsers.ps1 [-IncludeInactifs]
# ============================================================

[CmdletBinding()]
param(
    [switch]$IncludeInactifs = $false
)

# Culture FR — séparateur décimal virgule pour les exports CSV
[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES ET CHEMINS (identique aux autres scripts)
# ============================================================
$scriptName  = "Get-AssetsAccessUsers"
$scriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$secretsDir  = Join-Path $scriptRoot "secrets"
$exportsDir  = Join-Path $scriptRoot "exports"
$logsDir     = Join-Path $scriptRoot "logs"
$runStamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile     = Join-Path $logsDir ($scriptName + "_" + $runStamp + ".log")

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

foreach ($d in @($exportsDir, $secretsDir, $logsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# ============================================================
# 1. LOGGING ET HELPERS
# ============================================================
function Write-Log([string]$Message, [string]$Level = "INFO") {
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[" + $ts + "] [" + $Level + "] " + $Message
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch {}
}
function Write-Info([string]$Message) {
    Write-Host ("[INFO] " + $Message) -ForegroundColor Cyan
    Write-Log $Message "INFO"
}
function Write-Warn([string]$Message) {
    Write-Warning $Message
    Write-Log $Message "WARN"
}

# ============================================================
# 2. GESTION PROXY
# ============================================================
function Get-ProxyParams([string]$TargetUrl) {
    $params = @{}
    if ($ProxyUrl -and $ProxyUrl.Trim()) {
        $params.Proxy = $ProxyUrl
        if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true }
        return $params
    }
    if (-not $UseSystemProxy) { return $params }
    try {
        $dest = [uri]$TargetUrl
        $wp   = [System.Net.WebRequest]::DefaultWebProxy
        if ($wp -and -not $wp.IsBypassed($dest)) {
            $proxy = $wp.GetProxy($dest)
            if ($proxy -and $proxy.AbsoluteUri -ne $dest.AbsoluteUri) {
                $params.Proxy = $proxy.AbsoluteUri
                if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true }
            }
        }
    } catch {}
    return $params
}

function Invoke-ApiGet([string]$Url, [hashtable]$Headers) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-ProxyParams $Url; foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return ($raw | ConvertFrom-Json)
}

# ============================================================
# 3. CHARGEMENT CREDENTIALS JIRA
# ============================================================
Write-Info "=== 1/4 Chargement des credentials Jira ==="

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable : $jiraCredFile" }

$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

Write-Info ("Connecté sur : " + $jiraBaseUrl)
Write-Log ("=== DÉBUT EXECUTION Get-AssetsAccessUsers v1.2 — " + $runStamp + " ===") "INFO"

$usersAccessMap  = [ordered]@{}
$appAccountsSkipped = New-Object System.Collections.ArrayList

# ============================================================
# HELPER — Enregistrement utilisateur
# ============================================================
function Register-UserAccess(
    [string]$AccId,
    [string]$DisplayName,
    [string]$Email,
    [bool]$Active,
    [string]$AccountType,
    [string]$AccessSource,
    [string]$RoleDetail
) {
    if ([string]::IsNullOrWhiteSpace($AccId)) { return }
    if (-not $IncludeInactifs -and -not $Active) { return }

    if (-not $usersAccessMap.Contains($AccId)) {
        $usersAccessMap[$AccId] = [ordered]@{
            AccountId     = $AccId
            DisplayName   = $DisplayName
            EmailAddress  = $Email
            Active        = $Active
            AccountType   = $AccountType
            AccessSources = New-Object System.Collections.ArrayList
            RoleDetails   = New-Object System.Collections.ArrayList
        }
    }

    $u = $usersAccessMap[$AccId]
    if (-not [string]::IsNullOrWhiteSpace($DisplayName) -and [string]::IsNullOrWhiteSpace($u.DisplayName)) {
        $u.DisplayName  = $DisplayName
    }
    if (-not [string]::IsNullOrWhiteSpace($Email) -and [string]::IsNullOrWhiteSpace($u.EmailAddress)) {
        $u.EmailAddress = $Email
    }
    if (-not [string]::IsNullOrWhiteSpace($AccountType) -and [string]::IsNullOrWhiteSpace($u.AccountType)) {
        $u.AccountType  = $AccountType
    }

    if (-not $u.AccessSources.Contains($AccessSource)) {
        [void]$u.AccessSources.Add($AccessSource)
    }
    if (-not [string]::IsNullOrWhiteSpace($RoleDetail) -and -not $u.RoleDetails.Contains($RoleDetail)) {
        [void]$u.RoleDetails.Add($RoleDetail)
    }
}

# ============================================================
# 4. ÉTAPE 1 : NIVEAU PRODUIT & LICENCES JSM / ASSETS
# ============================================================
Write-Info "=== 2/4 Analyse des accès produit (Licences JSM & Groupes Admin) ==="

$targetGroups = New-Object System.Collections.Hashtable

# a) Recherche des groupes liés au rôle d'application JSM
try {
    $appRoleUrl = $jiraBaseUrl + "/rest/api/3/applicationrole/jira-servicedesk"
    $appRole    = Invoke-ApiGet -Url $appRoleUrl -Headers $jiraHeaders
    if ($appRole.groups) {
        foreach ($g in $appRole.groups) { $targetGroups[[string]$g] = "Licence JSM (Service Desk)" }
    }
} catch {
    Write-Warn "Rôle d'application jira-servicedesk non lu via API, utilisation des groupes standards."
}

# b) Ajout des groupes d'administration standards
$defaultAdminGroups = @("jira-servicedesk-users", "jira-administrators", "site-admins", "org-admins", "administrators")
foreach ($g in $defaultAdminGroups) {
    if (-not $targetGroups.ContainsKey($g)) {
        $targetGroups[$g] = "Groupe Produit / Admin Standard"
    }
}

Write-Info ("Nombre de groupes d'accès produit identifiés : " + $targetGroups.Count)

# c) Extraction des membres de ces groupes
foreach ($gName in $targetGroups.Keys) {
    $grpLabel = $targetGroups[$gName]
    $startAt  = 0
    $maxRes   = 50

    Write-Info ("Lecture des membres du groupe : " + $gName + "...")

    while ($true) {
        $memberUrl = $jiraBaseUrl + "/rest/api/3/group/member?groupname=" + [Uri]::EscapeDataString($gName) + "&startAt=" + $startAt + "&maxResults=" + $maxRes
        try {
            $respMembers = Invoke-ApiGet -Url $memberUrl -Headers $jiraHeaders
        } catch { break }

        if (-not $respMembers.values -or $respMembers.values.Count -eq 0) { break }

        foreach ($usr in $respMembers.values) {
            $accType = [string]$usr.accountType

            # Exclusion des comptes applicatifs (add-ons Marketplace, bots, comptes de service)
            if ($accType -ieq "app") {
                $msg = "Compte applicatif ignore : " + [string]$usr.displayName + " (accountType=app, groupe: " + $gName + ")"
                Write-Warn $msg
                if (-not $appAccountsSkipped.Contains([string]$usr.displayName)) {
                    [void]$appAccountsSkipped.Add([string]$usr.displayName)
                }
                continue
            }

            # Evaluation propre du booléen avant passage en paramètre
            $isAct = ($usr.active -eq $true -or [string]$usr.active -ieq "true")

            Register-UserAccess `
                -AccId        ([string]$usr.accountId) `
                -DisplayName  ([string]$usr.displayName) `
                -Email        ([string]$usr.emailAddress) `
                -Active       $isAct `
                -AccountType  $accType `
                -AccessSource "Licence JSM / Produit" `
                -RoleDetail   ("Groupe: " + $gName + " (" + $grpLabel + ")")
        }

        $total = if ($respMembers.total) { [int]$respMembers.total } else { 0 }
        if ($respMembers.isLast -eq $true -or ($startAt + $maxRes) -ge $total) { break }
        $startAt += $maxRes
    }
}

# ============================================================
# 5. ÉTAPE 2 : NIVEAU SCHÉMAS D'OBJETS ASSETS (Rôles In-Product)
# ============================================================
Write-Info "=== 3/4 Analyse des permissions sur les Schémas d'Objets Assets ==="

# Récupération du Workspace ID Assets
$wsUrl = $jiraBaseUrl + "/rest/servicedeskapi/assets/workspace"
$assetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1" # Default Fallback
try {
    $respWs = Invoke-ApiGet -Url $wsUrl -Headers $jiraHeaders
    if ($respWs.values -and $respWs.values.Count -gt 0) {
        $assetsWorkspaceId = [string]$respWs.values[0].workspaceId
    }
} catch {}

$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $assetsWorkspaceId + "/v1/objectschema/list"

try {
    $respSchemas = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
    $sList = if ($respSchemas.values) { $respSchemas.values } else { $respSchemas.objectSchemas }
    Write-Info ("Schémas d'objets Assets trouvés : " + $sList.Count)

    foreach ($sch in $sList) {
        $schId   = [string]$sch.id
        $schName = [string]$sch.name

        $configUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $assetsWorkspaceId + "/v1/objectschema/" + $schId
        try {
            $schDetails = Invoke-ApiGet -Url $configUrl -Headers $jiraHeaders
            if ($schDetails.roles) {
                foreach ($r in $schDetails.roles) {
                    $rName = [string]$r.name
                    if ($r.users) {
                        foreach ($u in $r.users) {
                            $accType = [string]$u.accountType
                            if ($accType -ieq "app") {
                                Write-Warn ("Compte applicatif ignore (schema Assets) : " + [string]$u.displayName)
                                continue
                            }
                            Register-UserAccess `
                                -AccId        ([string]$u.accountId) `
                                -DisplayName  ([string]$u.displayName) `
                                -Email        ([string]$u.emailAddress) `
                                -Active       $true `
                                -AccountType  $accType `
                                -AccessSource "Rôle Schéma Assets" `
                                -RoleDetail   ("Schéma: " + $schName + " | Rôle: " + $rName)
                        }
                    }
                }
            }
        } catch {}
    }
} catch {
    Write-Warn "Impossible d'interroger la liste des schémas d'objets Assets."
}

# ============================================================
# 6. CONSOLIDATION ET EXPORT CSV
# ============================================================
Write-Info "=== 4/4 Consolidation et Exportation CSV ==="

$finalRows = New-Object System.Collections.ArrayList

foreach ($accId in $usersAccessMap.Keys) {
    $u = $usersAccessMap[$accId]

    # Libellé lisible du type de compte
    $accTypeLib = switch ([string]$u.AccountType) {
        "atlassian" { "Utilisateur Humain" }
        "app"       { "Compte Applicatif" }
        "customer"  { "Client Portail JSM" }
        default     { [string]$u.AccountType }
    }

    [void]$finalRows.Add([pscustomobject]@{
        "NOM Prénom"          = $u.DisplayName
        "Adresse Email"       = $u.EmailAddress
        "Type de Compte"      = $accTypeLib
        "Statut Compte"       = if ($u.Active) { "Actif" } else { "Inactif" }
        "Sources d'Accès"     = ($u.AccessSources -join " ; ")
        "Détail des Rôles"    = ($u.RoleDetails -join " | ")
        "Identifiant Jira"    = $u.AccountId
    })
}

# Tri alphabétique par NOM Prénom
$finalRows = @($finalRows | Sort-Object "NOM Prénom")

$exportCsv = Join-Path $exportsDir ("Utilisateurs_Acces_Assets_Jiradot_" + $runStamp + ".csv")

# Export CSV au format Excel FR (UTF-8 BOM + séparateur ;)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$writer  = New-Object System.IO.StreamWriter($exportCsv, $false, $utf8Bom)
$headers = @("NOM Prénom","Adresse Email","Type de Compte","Statut Compte","Sources d'Accès","Détail des Rôles","Identifiant Jira")

$writer.WriteLine(($headers -join ";"))

foreach ($row in $finalRows) {
    $vals = New-Object System.Collections.ArrayList
    foreach ($h in $headers) {
        $s = [string]$row.$h
        if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) {
            [void]$vals.Add('"' + $s.Replace('"','""') + '"')
        } else { [void]$vals.Add($s) }
    }
    $writer.WriteLine(($vals -join ";"))
}
$writer.Close(); $writer.Dispose()

# Récapitulatif des comptes applicatifs ignorés
if ($appAccountsSkipped.Count -gt 0) {
    Write-Info ("--- Comptes applicatifs (add-ons) ignores : " + $appAccountsSkipped.Count + " ---")
    foreach ($appName in $appAccountsSkipped) {
        Write-Host ("  [APP] " + $appName) -ForegroundColor DarkYellow
    }
}

Write-Info ("Export terminé avec succès !")
Write-Info ("Nombre total d'utilisateurs humains ayant accès à Assets : " + $finalRows.Count)
Write-Info ("Fichier CSV généré dans : " + $exportCsv)
Write-Log  ("=== FIN EXECUTION Get-AssetsAccessUsers v1.2 — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ===") "INFO"

# Affichage du récapitulatif console
$finalRows | Select-Object -First 25 | Format-Table -Property "NOM Prénom","Adresse Email","Type de Compte","Statut Compte","Sources d'Accès" -AutoSize