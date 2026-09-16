<#
Detach des utilisateurs inactifs (AccountId) des rôles de projets Jira où ils sont acteurs (user actors)
- Input: CSV "uniques" (au moins colonne AccountId)
- TLS 1.2 + Proxy système Windows (PAC/WPAD/McAfee)
- Stockage local chiffré (DPAPI) de l'email+token (Export-Clixml)
- DryRun par défaut (ne supprime rien) -> utiliser -Execute pour appliquer

Exports:
  * C:\Temp\detach_inactive_users_actions.csv
  * C:\Temp\detach_inactive_users_summary.csv

Usage:
  .\detachInactiveUsersFromProjects.ps1
  .\detachInactiveUsersFromProjects.ps1 -Execute
  .\detachInactiveUsersFromProjects.ps1 -ResetCreds
#>

param(
    [string]$InputCsv,
    [switch]$Execute,
    [switch]$ResetCreds
)

# ----------------------------
# File picker (si InputCsv non fourni)
# ----------------------------
function Select-InputCsvFile([string]$initialDir) {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "Sélectionner le CSV des membres inactifs (uniques)"
        $dlg.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
        $dlg.Multiselect = $false
        if ($initialDir -and (Test-Path $initialDir)) { $dlg.InitialDirectory = $initialDir }

        $result = $dlg.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { throw "Aucun fichier sélectionné." }
        return $dlg.FileName
    } catch {
        Write-Warning "Boîte de dialogue indisponible, fallback en saisie manuelle."
        $p = Read-Host "Chemin complet du CSV"
        if (-not (Test-Path $p)) { throw "Fichier introuvable: $p" }
        return $p
    }
}

if ([string]::IsNullOrWhiteSpace($InputCsv)) {
    $defaultDir = "C:\Users\GUEDJ-F\OneDrive - Harmonie Mutuelle\Documents\powershell"
    $InputCsv = Select-InputCsvFile -initialDir $defaultDir
}
Write-Host "CSV sélectionné : $InputCsv" -ForegroundColor Cyan

# ----------------------------
# Runtime
# ----------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Proxy système Windows + creds Windows
$systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
$systemProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Get-ProxyForUrl($url) {
    try {
        $u = [Uri]$url
        $p = $systemProxy.GetProxy($u)
        if ($p -and $p.AbsoluteUri -ne $u.AbsoluteUri) { return $p.AbsoluteUri }
        return $null
    } catch { return $null }
}

function Invoke-JiraRest($method, $url, $headers, $body = $null) {
    $proxyUri = Get-ProxyForUrl $url

    $params = @{
        Method      = $method
        Uri         = $url
        Headers     = $headers
        ErrorAction = "Stop"
    }
    if ($null -ne $body) { $params.Body = $body }

    if ($proxyUri) {
        $params.Proxy = $proxyUri
        $params.ProxyUseDefaultCredentials = $true
    }

    try {
        return Invoke-RestMethod @params
    } catch {
        # enrichir l'erreur avec status code si dispo
        $status = $null
        try {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $status = [int]$_.Exception.Response.StatusCode
            }
        } catch { }

        $msg = $_.Exception.Message
        $details = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = $_.ErrorDetails.Message }
        if ($_.Exception.InnerException) { $msg = "$msg | Inner: $($_.Exception.InnerException.Message)" }

        if ($details) {
            if ($status) { throw "HTTP $status failed: $method $url => $msg | Details: $details" }
            throw "HTTP failed: $method $url => $msg | Details: $details"
        }
        if ($status) { throw "HTTP $status failed: $method $url => $msg" }
        throw "HTTP failed: $method $url => $msg"
    }
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportActionsPath = "C:\Temp\detach_inactive_users_actions.csv"
$exportSummaryPath = "C:\Temp\detach_inactive_users_summary.csv"

$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    if (-not $reset -and (Test-Path $credPath)) {
        try { return Import-Clixml -Path $credPath }
        catch { Write-Warning "Impossible de relire le fichier credential, nouvelle saisie requise. Détail: $_" }
    }

    $email = Read-Host "Email Atlassian (ex: prenom.nom@domaine.fr)"
    if ($email -match '^\[(.+?)\]\(mailto:(.+?)\)$') { $email = $Matches[2] }
    $email = ($email -replace '^mailto:', '').Trim()

    $secureToken = Read-Host "API Token Atlassian (saisie masquée)" -AsSecureString

    $cred = New-Object System.Management.Automation.PSCredential($email, $secureToken)
    $cred | Export-Clixml -Path $credPath
    Write-Host "Identifiants sauvegardés dans: $credPath" -ForegroundColor Green
    return $cred
}

# ----------------------------
# Auth header
# ----------------------------
$jiraCred = Get-JiraCredential -credPath $credPath -reset:$ResetCreds

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($jiraCred.Password)
try { $apiTokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

$pair   = "$($jiraCred.UserName)`:$apiTokenPlain"
$bytes  = [Text.Encoding]::UTF8.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)

$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
    "Content-Type"= "application/json"
}

# ----------------------------
# Test API
# ----------------------------
Write-Host "Test API /myself ..." -ForegroundColor Cyan
$me = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

Write-Host ("Mode: " + $(if ($Execute) { "EXECUTE (suppression réelle)" } else { "DRYRUN (aucune suppression)" })) -ForegroundColor Magenta

# ----------------------------
# Load inactive unique users CSV
# ----------------------------
Write-Host "Chargement des utilisateurs depuis CSV..." -ForegroundColor Cyan
$rows = Import-Csv -LiteralPath $InputCsv -Encoding UTF8

if (-not $rows -or $rows.Count -eq 0) { throw "CSV vide: $InputCsv" }

# Colonnes attendues (au moins AccountId)
$colNames = $rows[0].PSObject.Properties.Name
if (-not ($colNames -contains "AccountId")) {
    throw "Colonne 'AccountId' introuvable dans le CSV. Colonnes détectées: $($colNames -join ', ')"
}

# index users by accountId
$inactiveIndex = @{}
foreach ($r in $rows) {
    $aid = ($r.AccountId + "").Trim()
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }

    # Si le CSV contient Active, on filtre sur Active=false ; sinon on prend tout
    $keep = $true
    if ($colNames -contains "Active") {
        $a = ($r.Active + "").Trim().ToLowerInvariant()
        if ($a -eq "true" -or $a -eq "1") { $keep = $false }
    }

    if ($keep -and (-not $inactiveIndex.ContainsKey($aid))) {
        $inactiveIndex[$aid] = [PSCustomObject]@{
            AccountId    = $aid
            DisplayName  = $(if ($colNames -contains "DisplayName") { $r.DisplayName } else { $null })
            EmailAddress = $(if ($colNames -contains "EmailAddress") { $r.EmailAddress } else { $null })
            Groups       = $(if ($colNames -contains "Groups") { $r.Groups } else { $null })
        }
    }
}

Write-Host ("Nb comptes inactifs (cibles) : " + $inactiveIndex.Keys.Count) -ForegroundColor Green
if ($inactiveIndex.Keys.Count -eq 0) { throw "Aucun compte inactif cible trouvé (après filtrage éventuel Active=false)." }

# ----------------------------
# Get all live projects
# ----------------------------
Write-Host "Récupération des projets live..." -ForegroundColor Cyan
$allProjects = @()
$startAt = 0
$maxResults = 50
while ($true) {
    $url = "$siteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$maxResults&status=live"
    $resp = Invoke-JiraRest "GET" $url $headers
    if ($resp.values) { $allProjects += $resp.values }
    $startAt += ($resp.values | Measure-Object).Count
    if ($startAt -ge $resp.total) { break }
}
Write-Host ("Nombre de projets live: " + $allProjects.Count) -ForegroundColor Green

# ----------------------------
# Process roles & remove inactive user actors
# ----------------------------
$resultActions = @()
$nowIso = (Get-Date).ToString("s")

foreach ($p in $allProjects) {
    $projectKey  = $p.key
    $projectName = $p.name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    Write-Host "Projet $projectKey - $projectName" -ForegroundColor Yellow

    $rolesResp = $null
    try {
        $rolesResp = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey/role" $headers
    } catch {
        $resultActions += [PSCustomObject]@{
            Timestamp    = $nowIso
            Mode         = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
            Action       = "READ_ROLES_FAILED"
            Success      = $false
            ProjectKey   = $projectKey
            ProjectName  = $projectName
            RoleId       = $null
            RoleName     = $null
            AccountId    = $null
            DisplayName  = $null
            Details      = "$_"
        }
        continue
    }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl = $prop.Value

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRest "GET" $roleUrl $headers
        } catch {
            $resultActions += [PSCustomObject]@{
                Timestamp    = $nowIso
                Mode         = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
                Action       = "READ_ROLE_DETAIL_FAILED"
                Success      = $false
                ProjectKey   = $projectKey
                ProjectName  = $projectName
                RoleId       = $null
                RoleName     = $prop.Name
                AccountId    = $null
                DisplayName  = $null
                Details      = "$_"
            }
            continue
        }

        $roleId = $roleDetail.id
        $roleName = $roleDetail.name

        foreach ($actor in $roleDetail.actors) {

            # On ne cible que les acteurs "user" (avec accountId)
            $accountId = $null
            if ($actor.actorUser -and $actor.actorUser.accountId) { $accountId = $actor.actorUser.accountId }
            elseif ($actor.accountId) { $accountId = $actor.accountId }

            if ([string]::IsNullOrWhiteSpace($accountId)) { continue }

            if ($inactiveIndex.ContainsKey($accountId)) {

                $disp = $inactiveIndex[$accountId].DisplayName
                if ([string]::IsNullOrWhiteSpace($disp)) { $disp = $actor.displayName }

                # Log "found"
                $resultActions += [PSCustomObject]@{
                    Timestamp    = $nowIso
                    Mode         = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
                    Action       = "FOUND_IN_PROJECT_ROLE"
                    Success      = $true
                    ProjectKey   = $projectKey
                    ProjectName  = $projectName
                    RoleId       = $roleId
                    RoleName     = $roleName
                    AccountId    = $accountId
                    DisplayName  = $disp
                    Details      = "ActorType=$($actor.type)"
                }

                # Suppression
                if ($Execute) {
                    $delUrl = "$siteUrl/rest/api/3/project/$projectKey/role/$roleId?user=$([System.Uri]::EscapeDataString($accountId))"
                    try {
                        Invoke-JiraRest "DELETE" $delUrl $headers | Out-Null
                        $resultActions += [PSCustomObject]@{
                            Timestamp    = $nowIso
                            Mode         = "EXECUTE"
                            Action       = "REMOVE_USER_FROM_ROLE"
                            Success      = $true
                            ProjectKey   = $projectKey
                            ProjectName  = $projectName
                            RoleId       = $roleId
                            RoleName     = $roleName
                            AccountId    = $accountId
                            DisplayName  = $disp
                            Details      = "Deleted via $delUrl"
                        }
                    } catch {
                        $resultActions += [PSCustomObject]@{
                            Timestamp    = $nowIso
                            Mode         = "EXECUTE"
                            Action       = "REMOVE_USER_FROM_ROLE_FAILED"
                            Success      = $false
                            ProjectKey   = $projectKey
                            ProjectName  = $projectName
                            RoleId       = $roleId
                            RoleName     = $roleName
                            AccountId    = $accountId
                            DisplayName  = $disp
                            Details      = "$_"
                        }
                    }
                }
            }
        }
    }
}

# ----------------------------
# Summary per user
# ----------------------------
$summary = @()
if ($resultActions.Count -gt 0) {
    $summary = $resultActions |
        Where-Object { $_.Action -in @("FOUND_IN_PROJECT_ROLE","REMOVE_USER_FROM_ROLE","REMOVE_USER_FROM_ROLE_FAILED") } |
        Group-Object AccountId |
        ForEach-Object {
            $aid = $_.Name
            $one = $_.Group | Select-Object -First 1
            $foundProjects = ($_.Group | Where-Object Action -eq "FOUND_IN_PROJECT_ROLE" | Select-Object -ExpandProperty ProjectKey -Unique | Sort-Object) -join ";"
            $removedOk = ($_.Group | Where-Object Action -eq "REMOVE_USER_FROM_ROLE" | Measure-Object).Count
            $removedKo = ($_.Group | Where-Object Action -eq "REMOVE_USER_FROM_ROLE_FAILED" | Measure-Object).Count

            [PSCustomObject]@{
                AccountId           = $aid
                DisplayName         = $one.DisplayName
                ProjectsFoundKeys   = $foundProjects
                RemoveSuccessCount  = $removedOk
                RemoveFailedCount   = $removedKo
            }
        } | Sort-Object DisplayName
}

# ----------------------------
# Export CSV
# ----------------------------
$exportDir = Split-Path $exportActionsPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export actions : $exportActionsPath" -ForegroundColor Cyan
$resultActions | Export-Csv -Path $exportActionsPath -NoTypeInformation -Encoding UTF8

Write-Host "Export summary : $exportSummaryPath" -ForegroundColor Cyan
$summary | Export-Csv -Path $exportSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Actions : $exportActionsPath"
Write-Host " - Summary : $exportSummaryPath"
Write-Host " - Creds   : $credPath"