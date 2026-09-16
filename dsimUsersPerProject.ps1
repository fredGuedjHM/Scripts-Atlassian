<#
Liste des USERS ACTIFS provenant de groupes DSIM, projet par projet / rôle
- On se limite aux users membres des GROUPES dont le nom contient "DSIM"
- Ne remonte PAS les autres users (acteurs individuels, autres groupes)
- TLS 1.2 + Proxy système + DPAPI creds (email + API token)

Exports:
  * C:\Temp\dsim_actifs_par_projet_detail.csv
  * C:\Temp\dsim_actifs_par_projet_resume_user.csv
  * C:\Temp\dsim_actifs_par_projet_resume_projet.csv

Usage:
  .\dsimUsersPerProject.ps1
  .\dsimUsersPerProject.ps1 -Filter "DSIM"
  .\dsimUsersPerProject.ps1 -ResetCreds
#>

param(
    [string]$Filter = "DSIM",
    [switch]$ResetCreds
)

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

function Invoke-JiraRestWithRetry($method, $url, $headers, $body = $null, [int]$maxRetries = 5) {
    $attempt = 0
    while ($true) {
        try {
            return Invoke-JiraRest $method $url $headers $body
        } catch {
            $attempt++
            if ($_.Exception.Message -match '\bHTTP 429\b' -and $attempt -le $maxRetries) {
                $sleep = 30 + (5 * $attempt)
                Write-Warning "Rate limit (429). Retry dans $sleep sec (tentative $attempt/$maxRetries) : $url"
                Start-Sleep -Seconds $sleep
                continue
            }
            throw
        }
    }
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportDetailPath        = "C:\Temp\dsim_actifs_par_projet_detail.csv"
$exportUserSummaryPath   = "C:\Temp\dsim_actifs_par_projet_resume_user.csv"
$exportProjectSummaryPath= "C:\Temp\dsim_actifs_par_projet_resume_projet.csv"

$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    if (-not $reset -and (Test-Path $credPath)) {
        try { return Import-Clixml -Path $credPath }
        catch { Write-Warning "Impossible de relire le credential, nouvelle saisie requise. Détail: $_" }
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
$me = Invoke-JiraRestWithRetry "GET" "$siteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

# ----------------------------
# 1) Groupes DSIM
# ----------------------------
Write-Host "Recherche des groupes contenant: '$Filter' ..." -ForegroundColor Cyan
$encoded = [System.Uri]::EscapeDataString($Filter)
$groupsUrl = "$siteUrl/rest/api/3/groups/picker?query=$encoded&maxResults=1000&caseInsensitive=true"
$groupsResp = Invoke-JiraRestWithRetry "GET" $groupsUrl $headers

$groups = @()
if ($groupsResp.groups) { $groups = $groupsResp.groups }

$groups = $groups | Where-Object { $_.name -and ($_.name.ToUpper().Contains($Filter.ToUpper())) } | Sort-Object name

Write-Host ("Groupes DSIM trouvés: " + $groups.Count) -ForegroundColor Green
if ($groups.Count -eq 0) { throw "Aucun groupe ne correspond au filtre '$Filter'." }

# ----------------------------
# 2) Membres ACTIFS des groupes DSIM (index global par AccountId)
# ----------------------------
Write-Host "Construction de l'index des membres ACTIFS des groupes DSIM..." -ForegroundColor Cyan

$dsimUserIndex = @{}       # AccountId -> { DisplayName, Email, Groups (liste) }
$dsimUserGroups = @{}      # AccountId -> HashSet(GroupName)

foreach ($g in $groups) {
    $gName = $g.name
    $gId   = $g.groupId

    Write-Host "Groupe DSIM: $gName" -ForegroundColor Yellow

    $startAt = 0
    $maxResults = 50
    while ($true) {
        $url = $null
        if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $url = "$siteUrl/rest/api/3/group/member?groupId=$([System.Uri]::EscapeDataString($gId))&includeInactiveUsers=false&startAt=$startAt&maxResults=$maxResults"
        } else {
            $url = "$siteUrl/rest/api/3/group/member?groupname=$([System.Uri]::EscapeDataString($gName))&includeInactiveUsers=false&startAt=$startAt&maxResults=$maxResults"
        }

        $page = Invoke-JiraRestWithRetry "GET" $url $headers
        $values = @()
        if ($page.values) { $values = $page.values }

        foreach ($u in $values) {
            # includeInactiveUsers=false → on ne devrait avoir que des actifs, mais on vérifie quand même
            if ($u.active -ne $true) { continue }

            $aid = $u.accountId
            if ([string]::IsNullOrWhiteSpace($aid)) { continue }

            if (-not $dsimUserIndex.ContainsKey($aid)) {
                $mail = $null
                if ($u.PSObject.Properties.Name -contains "emailAddress") { $mail = $u.emailAddress }

                $dsimUserIndex[$aid] = [PSCustomObject]@{
                    AccountId    = $aid
                    DisplayName  = $u.displayName
                    EmailAddress = $mail
                }
                $dsimUserGroups[$aid] = New-Object System.Collections.Generic.HashSet[string]
            }
            [void]$dsimUserGroups[$aid].Add($gName)
        }

        $returned = ($values | Measure-Object).Count
        $startAt += $returned
        if ($returned -eq 0) { break }
        if ($page.isLast -eq $true) { break }
        if ($page.total -and $startAt -ge $page.total) { break }
    }
}

Write-Host ("Nb d'utilisateurs DSIM ACTIFS uniques : " + $dsimUserIndex.Keys.Count) -ForegroundColor Green
if ($dsimUserIndex.Keys.Count -eq 0) {
    Write-Warning "Aucun user actif trouvé dans les groupes DSIM (filter '$Filter')."
}

# ----------------------------
# 3) Projets live
# ----------------------------
Write-Host "Récupération des projets live..." -ForegroundColor Cyan
$allProjects = @()
$startAt = 0
$maxResults = 50
while ($true) {
    $url = "$siteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$maxResults&status=live"
    $resp = Invoke-JiraRestWithRetry "GET" $url $headers
    if ($resp.values) { $allProjects += $resp.values }
    $startAt += ($resp.values | Measure-Object).Count
    if ($startAt -ge $resp.total) { break }
}
Write-Host ("Nombre de projets live: " + $allProjects.Count) -ForegroundColor Green

# ----------------------------
# 4) Parcours projets / rôles / groupes DSIM
# ----------------------------
Write-Host "Analyse des rôles de projet pour les groupes DSIM..." -ForegroundColor Cyan

$resultDetail = @()

foreach ($p in $allProjects) {
    $projectKey  = $p.key
    $projectName = $p.name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    Write-Host "Projet $projectKey - $projectName" -ForegroundColor Yellow

    $rolesResp = $null
    try {
        $rolesResp = Invoke-JiraRestWithRetry "GET" "$siteUrl/rest/api/3/project/$projectKey/role" $headers
    } catch {
        Write-Warning "Impossible de lire les rôles pour $projectKey : $_"
        continue
    }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl  = $prop.Value

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRestWithRetry "GET" $roleUrl $headers
        } catch {
            Write-Warning "Détail rôle non récupéré ($($prop.Name)) pour $projectKey : $_"
            continue
        }

        $roleId   = $roleDetail.id
        $roleName = $roleDetail.name

        # On ne s'intéresse qu'aux ACTORS de type "group" dont le nom contient DSIM
        foreach ($actor in $roleDetail.actors) {
            if ($actor.type -ne "atlassian-group-role-actor") { continue }

            $groupNameInRole = $null
            if ($actor.actorGroup -and $actor.actorGroup.name) {
                $groupNameInRole = $actor.actorGroup.name
            } else {
                $groupNameInRole = $actor.displayName
            }

            if (-not $groupNameInRole) { continue }

            if (-not $groupNameInRole.ToUpper().Contains($Filter.ToUpper())) { continue }

            # Pour ce groupe DSIM, lister tous les users ACTIFS (via index)
            foreach ($kvp in $dsimUserIndex.GetEnumerator()) {
                $uInfo = $kvp.Value
                $aid   = $uInfo.AccountId

                # L'utilisateur doit être membre du groupe utilisé dans ce rôle
                if (-not $dsimUserGroups[$aid].Contains($groupNameInRole)) { continue }

                $resultDetail += [PSCustomObject]@{
                    "Project Key"      = $projectKey
                    "Project Name"     = $projectName
                    "Role Id"          = $roleId
                    "Role Name"        = $roleName
                    "DSIM Group Name"  = $groupNameInRole
                    "AccountId"        = $uInfo.AccountId
                    "DisplayName"      = $uInfo.DisplayName
                    "EmailAddress"     = $uInfo.EmailAddress
                }
            }
        }
    }
}

Write-Host ("Lignes détail (user DSIM actif x projet x rôle) : " + $resultDetail.Count) -ForegroundColor Green

# ----------------------------
# 5) Résumé par USER
# ----------------------------
$userSummary = @()
if ($resultDetail.Count -gt 0) {
    $userSummary = $resultDetail |
        Group-Object AccountId |
        ForEach-Object {
            $aid = $_.Name
            $one = $_.Group | Select-Object -First 1

            $projects = $_.Group |
                Select-Object -Property "Project Key", "Project Name" -Unique |
                Sort-Object "Project Key"

            $roles = $_.Group |
                Select-Object -Property "Role Name" -Unique |
                Sort-Object "Role Name"

            $groupsUsed = $_.Group |
                Select-Object -Property "DSIM Group Name" -Unique |
                Sort-Object "DSIM Group Name"

            [PSCustomObject]@{
                "AccountId"        = $aid
                "DisplayName"      = $one.DisplayName
                "EmailAddress"     = $one.EmailAddress
                "Projects Count"   = $projects.Count
                "Projects (Keys)"  = ($projects | ForEach-Object { $_."Project Key" }) -join ";"
                "Projects (Names)" = ($projects | ForEach-Object { $_."Project Name" }) -join ";"
                "Roles"            = ($roles | ForEach-Object { $_."Role Name" }) -join ";"
                "DSIM Groups"      = ($groupsUsed | ForEach-Object { $_."DSIM Group Name" }) -join ";"
            }
        } |
        Sort-Object "DisplayName"
}

# ----------------------------
# 6) Résumé par PROJET
# ----------------------------
$projectSummary = @()
if ($resultDetail.Count -gt 0) {
    $projectSummary = $resultDetail |
        Group-Object "Project Key" |
        ForEach-Object {
            $pKey = $_.Name
            $one  = $_.Group | Select-Object -First 1

            $users = $_.Group |
                Select-Object -Property "AccountId", "DisplayName" -Unique |
                Sort-Object "DisplayName"

            [PSCustomObject]@{
                "Project Key"     = $pKey
                "Project Name"    = $one."Project Name"
                "Users Count"     = $users.Count
                "Users (AccountId)" = ($users | ForEach-Object { $_."AccountId" }) -join ";"
                "Users (Names)"   = ($users | ForEach-Object { $_."DisplayName" }) -join ";"
            }
        } |
        Sort-Object "Project Key"
}

# ----------------------------
# 7) Export CSV
# ----------------------------
$exportDir = Split-Path $exportDetailPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export détail : $exportDetailPath" -ForegroundColor Cyan
$resultDetail | Export-Csv -Path $exportDetailPath -NoTypeInformation -Encoding UTF8

Write-Host "Export résumé par user : $exportUserSummaryPath" -ForegroundColor Cyan
$userSummary | Export-Csv -Path $exportUserSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Export résumé par projet : $exportProjectSummaryPath" -ForegroundColor Cyan
$projectSummary | Export-Csv -Path $exportProjectSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Détail      : $exportDetailPath"
Write-Host " - Résumé user : $exportUserSummaryPath"
Write-Host " - Résumé proj : $exportProjectSummaryPath"
Write-Host " - Creds       : $credPath"