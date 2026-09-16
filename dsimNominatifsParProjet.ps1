<#
DSIM nominatif par projet (Jira Cloud) - v3
- Liste les USERS DSIM (actifs) ajoutés NOMINATIVEMENT (user actors) dans les rôles de projets
- Objectif: supprimer ces user actors pour n'avoir que des groupes DSIM dans les rôles
- Correction v3: suppression via roleUrl (fourni par Jira) + ?user=accountId
- TLS 1.2 + Proxy Windows + creds DPAPI
- DryRun par défaut; -Execute pour supprimer

Exports:
  * C:\Temp\dsim_nominatifs_par_projet_detail.csv
  * C:\Temp\dsim_nominatifs_par_projet_resume_projet.csv
  * C:\Temp\dsim_nominatifs_par_projet_actions.csv

Usage:
  .\dsimNominatifsParProjet.ps1
  .\dsimNominatifsParProjet.ps1 -Execute
  .\dsimNominatifsParProjet.ps1 -Execute -ForceRemoveEvenIfNotCovered
  .\dsimNominatifsParProjet.ps1 -Filter "DSIM" -ResetCreds
#>

param(
    [string]$Filter = "DSIM",
    [switch]$Execute,
    [switch]$ForceRemoveEvenIfNotCovered,
    [switch]$ResetCreds
)

# ----------------------------
# Runtime
# ----------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
$systemProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Get-ProxyForUrl($url) {
    try {
        $uri = [Uri]$url
        $p = $systemProxy.GetProxy($uri)
        if ($p -and $p.AbsoluteUri -ne $uri.AbsoluteUri) { return $p.AbsoluteUri }
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

function Get-RoleIdFromRoleUrl([string]$roleUrl) {
    # attend .../rest/api/3/project/{KEY}/role/{ID}
    try {
        $u = [Uri]$roleUrl
        $last = $u.Segments[-1].TrimEnd('/')
        if ($last -match '^\d+$') { return [int]$last }
        return $null
    } catch {
        return $null
    }
}

function Is-ValidRoleUrl([string]$roleUrl) {
    # garde-fou : on ne delete que si URL finit par /role/<digits>
    return ($roleUrl -match '/role/\d+/?$')
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportDetailPath  = "C:\Temp\dsim_nominatifs_par_projet_detail.csv"
$exportProjectPath = "C:\Temp\dsim_nominatifs_par_projet_resume_projet.csv"
$exportActionsPath = "C:\Temp\dsim_nominatifs_par_projet_actions.csv"

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
Write-Host ("Mode: " + $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })) -ForegroundColor Magenta
if ($Execute -and $ForceRemoveEvenIfNotCovered) {
    Write-Warning "ForceRemoveEvenIfNotCovered=ON : risque de retirer un accès si aucun groupe DSIM ne couvre le rôle."
}

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
# 2) Index users DSIM ACTIFS + membership par groupe
# ----------------------------
Write-Host "Indexation des membres ACTIFS des groupes DSIM..." -ForegroundColor Cyan

$dsimUserIndex  = @{}  # accountId -> info
$groupMembers   = @{}  # groupName -> HashSet(accountId)

foreach ($g in $groups) {
    $gName = $g.name
    $gId   = $g.groupId
    $groupMembers[$gName] = New-Object System.Collections.Generic.HashSet[string]

    Write-Host "Groupe DSIM: $gName" -ForegroundColor Yellow

    $startAt = 0
    $maxResults = 50
    while ($true) {

        if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $url = "$siteUrl/rest/api/3/group/member?groupId=$([System.Uri]::EscapeDataString($gId))&includeInactiveUsers=false&startAt=$startAt&maxResults=$maxResults"
        } else {
            $url = "$siteUrl/rest/api/3/group/member?groupname=$([System.Uri]::EscapeDataString($gName))&includeInactiveUsers=false&startAt=$startAt&maxResults=$maxResults"
        }

        $page = Invoke-JiraRestWithRetry "GET" $url $headers
        $values = @()
        if ($page.values) { $values = $page.values }

        foreach ($u in $values) {
            if ($u.active -ne $true) { continue }
            $aid = $u.accountId
            if ([string]::IsNullOrWhiteSpace($aid)) { continue }

            [void]$groupMembers[$gName].Add($aid)

            if (-not $dsimUserIndex.ContainsKey($aid)) {
                $mail = $null
                if ($u.PSObject.Properties.Name -contains "emailAddress") { $mail = $u.emailAddress }

                $dsimUserIndex[$aid] = [PSCustomObject]@{
                    AccountId    = $aid
                    DisplayName  = $u.displayName
                    EmailAddress = $mail
                }
            }
        }

        $returned = ($values | Measure-Object).Count
        $startAt += $returned
        if ($returned -eq 0) { break }
        if ($page.isLast -eq $true) { break }
        if ($page.total -and $startAt -ge $page.total) { break }
    }
}

Write-Host ("Users DSIM actifs uniques indexés: " + $dsimUserIndex.Keys.Count) -ForegroundColor Green

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
# 4) Scan roles: DSIM nominatif uniquement
# ----------------------------
Write-Host "Scan des rôles projet: DSIM nominatif..." -ForegroundColor Cyan

$resultDetail = @()
$resultActions = @()
$ts = (Get-Date).ToString("s")

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
        # IMPORTANT: la value est l'URL du rôle (ex .../role/10002)
        $roleUrl = [string]$prop.Value

        if ([string]::IsNullOrWhiteSpace($roleUrl)) { continue }

        if (-not (Is-ValidRoleUrl $roleUrl)) {
            # garde-fou: ne jamais supprimer si l'URL ne finit pas par /role/<digits>
            $resultActions += [PSCustomObject]@{
                Timestamp   = $ts
                Mode        = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
                Action      = "SKIP_INVALID_ROLEURL"
                Success     = $true
                ProjectKey  = $projectKey
                ProjectName = $projectName
                RoleId      = $null
                RoleName    = $prop.Name
                AccountId   = $null
                DisplayName = $null
                CoveredByGroupInSameRole = $null
                Details     = "roleUrl='$roleUrl'"
            }
            continue
        }

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRestWithRetry "GET" $roleUrl $headers
        } catch {
            Write-Warning "Détail rôle non récupéré ($($prop.Name)) pour $projectKey : $_"
            continue
        }

        $roleId = Get-RoleIdFromRoleUrl $roleUrl
        $roleName = $roleDetail.name

        # groupes DSIM présents dans CE rôle
        $dsimGroupsInRole = New-Object System.Collections.Generic.HashSet[string]
        foreach ($a in $roleDetail.actors) {
            if ($a.type -ne "atlassian-group-role-actor") { continue }
            $gn = $null
            if ($a.actorGroup -and $a.actorGroup.name) { $gn = $a.actorGroup.name } else { $gn = $a.displayName }
            if (-not $gn) { continue }
            if ($gn.ToUpper().Contains($Filter.ToUpper())) { [void]$dsimGroupsInRole.Add($gn) }
        }

        foreach ($a in $roleDetail.actors) {

            # acteur user ?
            $accountId = $null
            if ($a.actorUser -and $a.actorUser.accountId) { $accountId = $a.actorUser.accountId }
            elseif ($a.accountId) { $accountId = $a.accountId }

            if ([string]::IsNullOrWhiteSpace($accountId)) { continue }
            if (-not $dsimUserIndex.ContainsKey($accountId)) { continue }  # uniquement DSIM

            $uInfo = $dsimUserIndex[$accountId]

            # couverture: user membre d'un groupe DSIM présent dans le même rôle ?
            $covered = $false
            $coveringGroups = @()
            foreach ($gname in $dsimGroupsInRole) {
                if ($groupMembers.ContainsKey($gname) -and $groupMembers[$gname].Contains($accountId)) {
                    $covered = $true
                    $coveringGroups += $gname
                }
            }

            $resultDetail += [PSCustomObject]@{
                "Project Key"                  = $projectKey
                "Project Name"                 = $projectName
                "Role Url"                     = $roleUrl
                "Role Id"                      = $roleId
                "Role Name"                    = $roleName
                "AccountId"                    = $uInfo.AccountId
                "DisplayName"                  = $uInfo.DisplayName
                "EmailAddress"                 = $uInfo.EmailAddress
                "Nominative Actor Type"        = $a.type
                "DSIM Groups In Same Role"     = ($dsimGroupsInRole | Sort-Object) -join ";"
                "Covered By DSIM Group (Role)" = $covered
                "Covering DSIM Groups"         = ($coveringGroups | Sort-Object) -join ";"
            }

            # URL DELETE: on utilise roleUrl directement (robuste)
            $encodedUser = [System.Uri]::EscapeDataString($uInfo.AccountId)
            $deleteUrl = "$($roleUrl.TrimEnd('/'))?user=$encodedUser"

            if (-not $Execute) {
                $resultActions += [PSCustomObject]@{
                    Timestamp   = $ts
                    Mode        = "DRYRUN"
                    Action      = "WOULD_REMOVE_NOMINATIVE_USER_ACTOR"
                    Success     = $true
                    ProjectKey  = $projectKey
                    ProjectName = $projectName
                    RoleId      = $roleId
                    RoleName    = $roleName
                    AccountId   = $uInfo.AccountId
                    DisplayName = $uInfo.DisplayName
                    CoveredByGroupInSameRole = $covered
                    Details     = $deleteUrl
                }
            } else {
                if ($covered -or $ForceRemoveEvenIfNotCovered) {
                    try {
                        Invoke-JiraRestWithRetry "DELETE" $deleteUrl $headers | Out-Null
                        $resultActions += [PSCustomObject]@{
                            Timestamp   = $ts
                            Mode        = "EXECUTE"
                            Action      = "REMOVE_NOMINATIVE_USER_ACTOR"
                            Success     = $true
                            ProjectKey  = $projectKey
                            ProjectName = $projectName
                            RoleId      = $roleId
                            RoleName    = $roleName
                            AccountId   = $uInfo.AccountId
                            DisplayName = $uInfo.DisplayName
                            CoveredByGroupInSameRole = $covered
                            Details     = $deleteUrl
                        }
                    } catch {
                        $resultActions += [PSCustomObject]@{
                            Timestamp   = $ts
                            Mode        = "EXECUTE"
                            Action      = "REMOVE_NOMINATIVE_USER_ACTOR_FAILED"
                            Success     = $false
                            ProjectKey  = $projectKey
                            ProjectName = $projectName
                            RoleId      = $roleId
                            RoleName    = $roleName
                            AccountId   = $uInfo.AccountId
                            DisplayName = $uInfo.DisplayName
                            CoveredByGroupInSameRole = $covered
                            Details     = "$_ | deleteUrl=$deleteUrl"
                        }
                    }
                } else {
                    $resultActions += [PSCustomObject]@{
                        Timestamp   = $ts
                        Mode        = "EXECUTE"
                        Action      = "SKIP_NOT_COVERED_BY_DSIM_GROUP_IN_SAME_ROLE"
                        Success     = $true
                        ProjectKey  = $projectKey
                        ProjectName = $projectName
                        RoleId      = $roleId
                        RoleName    = $roleName
                        AccountId   = $uInfo.AccountId
                        DisplayName = $uInfo.DisplayName
                        CoveredByGroupInSameRole = $covered
                        Details     = "Skipped (not covered). deleteUrl would be: $deleteUrl"
                    }
                }
            }
        }
    }
}

Write-Host ("Lignes DSIM nominatif trouvées: " + $resultDetail.Count) -ForegroundColor Green

# ----------------------------
# Résumé par projet
# ----------------------------
$projectSummary = @()
if ($resultDetail.Count -gt 0) {
    $projectSummary = $resultDetail |
        Group-Object "Project Key" |
        ForEach-Object {
            $pKey = $_.Name
            $one  = $_.Group | Select-Object -First 1
            $users = $_.Group | Select-Object AccountId, DisplayName -Unique | Sort-Object DisplayName

            [PSCustomObject]@{
                "Project Key"   = $pKey
                "Project Name"  = $one."Project Name"
                "Nominative DSIM Users Count" = $users.Count
                "Users (Names)" = ($users | ForEach-Object { $_.DisplayName }) -join ";"
                "Users (AccountId)" = ($users | ForEach-Object { $_.AccountId }) -join ";"
            }
        } | Sort-Object "Project Key"
}

# ----------------------------
# Export
# ----------------------------
$exportDir = Split-Path $exportDetailPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export détail  : $exportDetailPath" -ForegroundColor Cyan
$resultDetail | Export-Csv -Path $exportDetailPath -NoTypeInformation -Encoding UTF8

Write-Host "Export résumé projet : $exportProjectPath" -ForegroundColor Cyan
$projectSummary | Export-Csv -Path $exportProjectPath -NoTypeInformation -Encoding UTF8

Write-Host "Export actions : $exportActionsPath" -ForegroundColor Cyan
$resultActions | Export-Csv -Path $exportActionsPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Détail  : $exportDetailPath"
Write-Host " - Résumé  : $exportProjectPath"
Write-Host " - Actions : $exportActionsPath"
Write-Host " - Creds   : $credPath"