# ============================================================
# Nom          : Get-JiraSuperAdmins.ps1
# Description  : Liste exhaustive des super-administrateurs
#                et administrateurs de l'instance Jiradot.
#                Interroge dynamiquement les groupes d'admin
#                via l'API Jira REST v3, consolide les membres
#                et exporte le resultat en CSV (Excel FR).
# Version      : 1.2
# Date         : 11/09/2026
# Auteur       : DSIM / HM_DSIM_PACT
# Prerequis    : Fichier secrets/jira-jiradot.cred.xml valide
# Usage        : .\Get-JiraSuperAdmins.ps1 [-IncludeInactive]
# ============================================================

[CmdletBinding()]
param(
    [switch]$IncludeInactive = $false
)

# Culture FR — séparateur décimal virgule pour les exports CSV
[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES ET CHEMINS (identique à Invoke-EtatsFinanciers)
# ============================================================
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$secretsDir  = Join-Path $scriptRoot "secrets"
$exportsDir  = Join-Path $scriptRoot "exports"
$logsDir     = Join-Path $scriptRoot "logs"
$runStamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile     = Join-Path $logsDir ("Get-JiraSuperAdmins_" + $runStamp + ".log")

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

foreach ($d in @($exportsDir, $logsDir)) {
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
# 2. GESTION PROXY (identique à Invoke-EtatsFinanciers)
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

# ============================================================
# 3. HELPER API GET (identique à Invoke-EtatsFinanciers)
# ============================================================
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
# 4. CHARGEMENT CREDENTIALS JIRA (identique à Invoke-EtatsFinanciers)
# ============================================================
Write-Info "=== 1/3 Chargement credentials Jira ==="

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable : $jiraCredFile" }

$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

Write-Info ("Connecte sur : " + $jiraBaseUrl)
Write-Log ("=== DEBUT EXECUTION Get-JiraSuperAdmins v1.2 — " + $runStamp + " ===") "INFO"

# ============================================================
# 5. RECHERCHE ET LISTAGE DES GROUPES ADMINS
# ============================================================
Write-Info "=== 2/3 Identification et extraction des groupes d'administration ==="

$groupsPickerUrl = $jiraBaseUrl + "/rest/api/3/groups/picker?query=admin&maxResults=50"
$respPicker      = Invoke-ApiGet -Url $groupsPickerUrl -Headers $jiraHeaders

$targetGroups = New-Object System.Collections.ArrayList
if ($respPicker.groups) {
    foreach ($grp in $respPicker.groups) {
        [void]$targetGroups.Add([pscustomobject]@{ Name = [string]$grp.name; GroupId = [string]$grp.groupId })
    }
}

Write-Info ("Groupes d'administration identifies : " + $targetGroups.Count)
foreach ($g in $targetGroups) {
    Write-Host ("  - " + $g.Name + " (ID: " + $g.GroupId + ")") -ForegroundColor Gray
}

# ============================================================
# 6. EXTRACTION DES MEMBRES PAR GROUPE (AVEC PAGINATION)
# ============================================================
$adminUsersMap = [ordered]@{}

foreach ($grp in $targetGroups) {
    $gName   = $grp.Name
    $gId     = $grp.GroupId
    $startAt = 0
    $maxRes  = 50

    Write-Info ("Lecture des membres du groupe : " + $gName + "...")

    while ($true) {
        $memberUrl = if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $jiraBaseUrl + "/rest/api/3/group/member?groupId=" + $gId + "&startAt=" + $startAt + "&maxResults=" + $maxRes
        } else {
            $jiraBaseUrl + "/rest/api/3/group/member?groupname=" + [Uri]::EscapeDataString($gName) + "&startAt=" + $startAt + "&maxResults=" + $maxRes
        }

        try {
            $respMembers = Invoke-ApiGet -Url $memberUrl -Headers $jiraHeaders
        } catch {
            Write-Warn ("Impossible de recuperer les membres du groupe " + $gName + " : " + $_.Exception.Message)
            break
        }

        if (-not $respMembers.values -or $respMembers.values.Count -eq 0) { break }

        foreach ($usr in $respMembers.values) {
            $accId = [string]$usr.accountId
            if ([string]::IsNullOrWhiteSpace($accId)) { continue }

            # Filtrage comptes inactifs sauf si -IncludeInactive
            if (-not $IncludeInactive -and $usr.active -eq $false) { continue }

            if (-not $adminUsersMap.Contains($accId)) {
                $adminUsersMap[$accId] = [ordered]@{
                    AccountId    = $accId
                    DisplayName  = [string]$usr.displayName
                    EmailAddress = [string]$usr.emailAddress
                    Active       = [bool]$usr.active
                    AccountType  = [string]$usr.accountType
                    AdminGroups  = New-Object System.Collections.ArrayList
                }
            }
            if (-not $adminUsersMap[$accId].AdminGroups.Contains($gName)) {
                [void]$adminUsersMap[$accId].AdminGroups.Add($gName)
            }
        }

        $total = if ($respMembers.total) { [int]$respMembers.total } else { 0 }
        if ($respMembers.isLast -eq $true -or ($startAt + $maxRes) -ge $total) { break }
        $startAt += $maxRes
    }
}

# ============================================================
# 7. CONSOLIDATION ET EXPORT CSV
# ============================================================
Write-Info "=== 3/3 Consolidation et Export CSV ==="

$finalRows = New-Object System.Collections.ArrayList

foreach ($accId in $adminUsersMap.Keys) {
    $u = $adminUsersMap[$accId]
    [void]$finalRows.Add([pscustomobject]@{
        "Nom Prenom"             = $u.DisplayName
        "Adresse Email"          = $u.EmailAddress
        "Statut Compte"          = if ($u.Active) { "Actif" } else { "Inactif" }
        "Type de Compte"         = $u.AccountType
        "Groupes Admin Detenus"  = ($u.AdminGroups -join " ; ")
        "Identifiant Jira"       = $u.AccountId
    })
}

# Tri alphabétique par Nom Prénom
$finalRows = @($finalRows | Sort-Object "Nom Prenom")

$exportCsv = Join-Path $exportsDir ("SuperAdministrateurs_Jiradot_" + $runStamp + ".csv")

# Export CSV avec la fonction maison (séparateur ; + UTF8 BOM + locale FR)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$writer  = New-Object System.IO.StreamWriter($exportCsv, $false, $utf8Bom)
$headers = @("Nom Prenom","Adresse Email","Statut Compte","Type de Compte","Groupes Admin Detenus","Identifiant Jira")
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

Write-Info ("Export termine avec succes !")
Write-Info ("Nombre total d'administrateurs identifies : " + $finalRows.Count)
Write-Info ("Fichier CSV genere dans : " + $exportCsv)
Write-Log  ("=== FIN EXECUTION Get-JiraSuperAdmins v1.2 — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ===") "INFO"

# Affichage récapitulatif console
$finalRows | Format-Table -Property "Nom Prenom","Adresse Email","Statut Compte","Groupes Admin Detenus" -AutoSize