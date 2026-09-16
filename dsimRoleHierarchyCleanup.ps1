<#
Pré-nettoyage DSIM par projet : suppression des rôles nominatif inférieurs selon hiérarchie
- Règles:
    * Si user DSIM a rôle Manager (R1), retirer nominatif du rôle User (R2) et ReadOnly (R3)
    * Si user DSIM a rôle User (R2), retirer nominatif du rôle ReadOnly (R3)
- Important: ne supprime QUE les affectations NOMINATIVES (actor user). Si l'affectation vient d'un groupe, on log et on skip.
- Détection DSIM: users appartenant à au moins un groupe dont le nom contient "DSIM" (filtre paramétrable)

DryRun par défaut, -Execute pour appliquer.

Exports:
  * C:\Temp\dsim_role_hierarchy_found.csv
  * C:\Temp\dsim_role_hierarchy_actions.csv
  * C:\Temp\dsim_role_hierarchy_summary.csv

Usage:
  .\dsimRoleHierarchyCleanup.ps1
  .\dsimRoleHierarchyCleanup.ps1 -Execute
  .\dsimRoleHierarchyCleanup.ps1 -RoleManager "1. Gestionnaire Projet" -RoleUser "2- utilisateur" -RoleReadOnly "3- lecture seule"
  .\dsimRoleHierarchyCleanup.ps1 -ResetCreds
#>

param(
    [string]$Filter = "DSIM",

    # Noms exacts (ou quasi exacts) des rôles à traiter
    [string]$RoleManager  = "1. Gestionnaire Projet",
    [string]$RoleUser     = "2- utilisateur",
    [string]$RoleReadOnly = "3- lecture seule",

    [switch]$Execute,
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

function Is-ValidRoleUrl([string]$roleUrl) {
    return ($roleUrl -match '/role/\d+/?$')
}

function Normalize-RoleName([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return ($s.Trim().ToLowerInvariant() -replace '\s+', ' ')
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportFoundPath   = "C:\Temp\dsim_role_hierarchy_found.csv"
$exportActionsPath = "C:\Temp\dsim_role_hierarchy_actions.csv"
$exportSummaryPath = "C:\Temp\dsim_role_hierarchy_summary.csv"

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

# ----------------------------
# 1) Groupes DSIM + membres actifs => index DSIM
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

# dsimUserIndex: accountId -> {DisplayName, Email}
$dsimUserIndex = @{}
# groupMembers: groupName -> HashSet(accountId)
$groupMembers = @{}

Write-Host "Indexation des membres ACTIFS des groupes DSIM..." -ForegroundColor Cyan

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
# 2) Projets live
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
# 3) Analyse par projet, calcul des overlaps, suppression nominatif des rôles inférieurs
# ----------------------------
$r1 = Normalize-RoleName $RoleManager
$r2 = Normalize-RoleName $RoleUser
$r3 = Normalize-RoleName $RoleReadOnly

$resultFound   = @()  # audit des présences (DSIM only)
$resultActions = @()  # ce qu'on supprime / ce qu'on ne peut pas
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

    # Pour ce projet, on veut récupérer roleUrl + roleDetail pour nos 3 rôles cibles
    $roleInfoByNormalizedName = @{}  # normName -> @{ roleUrl, roleId, roleName, actors }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl = [string]$prop.Value
        if ([string]::IsNullOrWhiteSpace($roleUrl)) { continue }
        if (-not (Is-ValidRoleUrl $roleUrl)) { continue }

        $detail = $null
        try { $detail = Invoke-JiraRestWithRetry "GET" $roleUrl $headers }
        catch { continue }

        $norm = Normalize-RoleName $detail.name
        if ($norm -in @($r1,$r2,$r3)) {
            $roleInfoByNormalizedName[$norm] = @{
                roleUrl  = $roleUrl.TrimEnd('/')
                roleName = $detail.name
                actors   = $detail.actors
            }
        }
    }

    # Si aucun des 3 rôles n'existe sur ce projet, on skip
    if ($roleInfoByNormalizedName.Keys.Count -eq 0) { continue }

    # Fonctions locales pour calculer membership DSIM d'un rôle
    function Get-DsimUsersInRoleEffective($actors) {
        $set = New-Object System.Collections.Generic.HashSet[string]

        foreach ($a in $actors) {
            # user nominatif ?
            $aid = $null
            if ($a.actorUser -and $a.actorUser.accountId) { $aid = $a.actorUser.accountId }
            elseif ($a.accountId) { $aid = $a.accountId }

            if ($aid -and $dsimUserIndex.ContainsKey($aid)) {
                [void]$set.Add($aid)
            }

            # groupes DSIM ?
            if ($a.type -eq "atlassian-group-role-actor") {
                $gn = $null
                if ($a.actorGroup -and $a.actorGroup.name) { $gn = $a.actorGroup.name } else { $gn = $a.displayName }
                if ($gn -and $gn.ToUpper().Contains($Filter.ToUpper())) {
                    # On ne garde que les users DSIM: membres de ce groupe
                    if ($groupMembers.ContainsKey($gn)) {
                        foreach ($m in $groupMembers[$gn]) { [void]$set.Add($m) }
                    }
                }
            }
        }
        return $set
    }

    function Get-DsimUsersInRoleNominativeOnly($actors) {
        $set = New-Object System.Collections.Generic.HashSet[string]
        foreach ($a in $actors) {
            $aid = $null
            if ($a.actorUser -and $a.actorUser.accountId) { $aid = $a.actorUser.accountId }
            elseif ($a.accountId) { $aid = $a.accountId }
            if ($aid -and $dsimUserIndex.ContainsKey($aid)) { [void]$set.Add($aid) }
        }
        return $set
    }

    # sets effectifs
    $R1eff = New-Object System.Collections.Generic.HashSet[string]
    $R2eff = New-Object System.Collections.Generic.HashSet[string]
    $R3eff = New-Object System.Collections.Generic.HashSet[string]

    # sets nominatif (ce qu'on peut supprimer)
    $R2nom = New-Object System.Collections.Generic.HashSet[string]
    $R3nom = New-Object System.Collections.Generic.HashSet[string]

    if ($roleInfoByNormalizedName.ContainsKey($r1)) {
        $R1eff = Get-DsimUsersInRoleEffective $roleInfoByNormalizedName[$r1].actors
    }
    if ($roleInfoByNormalizedName.ContainsKey($r2)) {
        $R2eff = Get-DsimUsersInRoleEffective $roleInfoByNormalizedName[$r2].actors
        $R2nom = Get-DsimUsersInRoleNominativeOnly $roleInfoByNormalizedName[$r2].actors
    }
    if ($roleInfoByNormalizedName.ContainsKey($r3)) {
        $R3eff = Get-DsimUsersInRoleEffective $roleInfoByNormalizedName[$r3].actors
        $R3nom = Get-DsimUsersInRoleNominativeOnly $roleInfoByNormalizedName[$r3].actors
    }

    # Audit (présence effective par rôle)
    foreach ($aid in $R1eff) {
        $u = $dsimUserIndex[$aid]
        $resultFound += [PSCustomObject]@{
            "Project Key" = $projectKey
            "Project Name"= $projectName
            "Role"        = $RoleManager
            "AccountId"   = $aid
            "DisplayName" = $u.DisplayName
        }
    }
    foreach ($aid in $R2eff) {
        $u = $dsimUserIndex[$aid]
        $resultFound += [PSCustomObject]@{
            "Project Key" = $projectKey
            "Project Name"= $projectName
            "Role"        = $RoleUser
            "AccountId"   = $aid
            "DisplayName" = $u.DisplayName
        }
    }
    foreach ($aid in $R3eff) {
        $u = $dsimUserIndex[$aid]
        $resultFound += [PSCustomObject]@{
            "Project Key" = $projectKey
            "Project Name"= $projectName
            "Role"        = $RoleReadOnly
            "AccountId"   = $aid
            "DisplayName" = $u.DisplayName
        }
    }

    # Règles d'overlap => à retirer nominativement:
    # - si R1 et R2: retirer R2 (nominatif seulement)
    # - si R1 et R3: retirer R3 (nominatif seulement)
    # - si R2 et R3: retirer R3 (nominatif seulement)
    $toRemoveR2 = New-Object System.Collections.Generic.HashSet[string]
    $toRemoveR3 = New-Object System.Collections.Generic.HashSet[string]

    foreach ($aid in $R2eff) { if ($R1eff.Contains($aid)) { [void]$toRemoveR2.Add($aid) } }
    foreach ($aid in $R3eff) {
        if ($R1eff.Contains($aid) -or $R2eff.Contains($aid)) { [void]$toRemoveR3.Add($aid) }
    }

    # Application: suppression uniquement si l'user est NOMINATIF dans le rôle cible
    if ($toRemoveR2.Count -gt 0 -and $roleInfoByNormalizedName.ContainsKey($r2)) {
        $roleUrl = $roleInfoByNormalizedName[$r2].roleUrl
        foreach ($aid in $toRemoveR2) {
            $u = $dsimUserIndex[$aid]

            if (-not $R2nom.Contains($aid)) {
                $resultActions += [PSCustomObject]@{
                    Timestamp   = $ts
                    Mode        = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
                    Action      = "CANNOT_REMOVE_R2_GROUP_BASED_OR_NOT_NOMINATIVE"
                    Success     = $false
                    ProjectKey  = $projectKey
                    ProjectName = $projectName
                    FromRole    = $RoleUser
                    Reason      = "User DSIM a aussi '$RoleManager' ; mais '$RoleUser' vient d'un groupe ou n'est pas nominatif."
                    AccountId   = $aid
                    DisplayName = $u.DisplayName
                    Details     = $roleUrl
                }
                continue
            }

            $deleteUrl = "$roleUrl?user=$([System.Uri]::EscapeDataString($aid))"

            if (-not $Execute) {
                $resultActions += [PSCustomObject]@{
                    Timestamp   = $ts
                    Mode        = "DRYRUN"
                    Action      = "WOULD_REMOVE_NOMINATIVE_FROM_R2"
                    Success     = $true
                    ProjectKey  = $projectKey
                    ProjectName = $projectName
                    FromRole    = $RoleUser
                    Reason      = "Overlap: a déjà '$RoleManager'"
                    AccountId   = $aid
                    DisplayName = $u.DisplayName
                    Details     = $deleteUrl
                }
            } else {
                try {
                    Invoke-JiraRestWithRetry "DELETE" $deleteUrl $headers | Out-Null
                    $resultActions += [PSCustomObject]@{
                        Timestamp   = $ts
                        Mode        = "EXECUTE"
                        Action      = "REMOVE_NOMINATIVE_FROM_R2"
                        Success     = $true
                        ProjectKey  = $projectKey
                        ProjectName = $projectName
                        FromRole    = $RoleUser
                        Reason      = "Overlap: a déjà '$RoleManager'"
                        AccountId   = $aid
                        DisplayName = $u.DisplayName
                        Details     = $deleteUrl
                    }
                } catch {
                    $resultActions += [PSCustomObject]@{
                        Timestamp   = $ts
                        Mode        = "EXECUTE"
                        Action      = "REMOVE_NOMINATIVE_FROM_R2_FAILED"
                        Success     = $false
                        ProjectKey  = $projectKey
                        ProjectName = $projectName
                        FromRole    = $RoleUser
                        Reason      = "Overlap: a déjà '$RoleManager'"
                        AccountId   = $aid
                        DisplayName = $u.DisplayName
                        Details     = "$_ | deleteUrl=$deleteUrl"
                    }
                }
            }
        }
    }

    if ($toRemoveR3.Count -gt 0 -and $roleInfoByNormalizedName.ContainsKey($r3)) {
        $roleUrl = $roleInfoByNormalizedName[$r3].roleUrl
        foreach ($aid in $toRemoveR3) {
            $u = $dsimUserIndex[$aid]

            if (-not $R3nom.Contains($aid)) {
                $resultActions += [PSCustomObject]@{
                    Timestamp   = $ts
                    Mode        = $(if ($Execute) { "EXECUTE" } else { "DRYRUN" })
                    Action      = "CANNOT_REMOVE_R3_GROUP_BASED_OR_NOT_NOMINATIVE"
                    Success     = $false
                    ProjectKey  = $projectKey
                    ProjectName = $projectName
                    FromRole    = $RoleReadOnly
                    Reason      = "User DSIM a un rôle supérieur (R1 ou R2) ; mais R3 vient d'un groupe ou n'est pas nominatif."
                    AccountId   = $aid
                    DisplayName = $u.DisplayName
                    Details     = $roleUrl
                }
                continue
            }

            $deleteUrl = "$roleUrl?user=$([System.Uri]::EscapeDataString($aid))"

            $why = $(if ($R1eff.Contains($aid)) { "Overlap: a déjà '$RoleManager'" } else { "Overlap: a déjà '$RoleUser'" })

            if (-not $Execute) {
                $resultActions += [PSCustomObject]@{
                    Timestamp   = $ts
                    Mode        = "DRYRUN"
                    Action      = "WOULD_REMOVE_NOMINATIVE_FROM_R3"
                    Success     = $true
                    ProjectKey  = $projectKey
                    ProjectName = $projectName
                    FromRole    = $RoleReadOnly
                    Reason      = $why
                    AccountId   = $aid
                    DisplayName = $u.DisplayName
                    Details     = $deleteUrl
                }
            } else {
                try {
                    Invoke-JiraRestWithRetry "DELETE" $deleteUrl $headers | Out-Null
                    $resultActions += [PSCustomObject]@{
                        Timestamp   = $ts
                        Mode        = "EXECUTE"
                        Action      = "REMOVE_NOMINATIVE_FROM_R3"
                        Success     = $true
                        ProjectKey  = $projectKey
                        ProjectName = $projectName
                        FromRole    = $RoleReadOnly
                        Reason      = $why
                        AccountId   = $aid
                        DisplayName = $u.DisplayName
                        Details     = $deleteUrl
                    }
                } catch {
                    $resultActions += [PSCustomObject]@{
                        Timestamp   = $ts
                        Mode        = "EXECUTE"
                        Action      = "REMOVE_NOMINATIVE_FROM_R3_FAILED"
                        Success     = $false
                        ProjectKey  = $projectKey
                        ProjectName = $projectName
                        FromRole    = $RoleReadOnly
                        Reason      = $why
                        AccountId   = $aid
                        DisplayName = $u.DisplayName
                        Details     = "$_ | deleteUrl=$deleteUrl"
                    }
                }
            }
        }
    }
}

# ----------------------------
# Summary
# ----------------------------
$summary = @()
if ($resultActions.Count -gt 0) {
    $summary = $resultActions |
        Group-Object ProjectKey |
        ForEach-Object {
            $one = $_.Group | Select-Object -First 1
            $ok  = ($_.Group | Where-Object Success -eq $true | Measure-Object).Count
            $ko  = ($_.Group | Where-Object Success -eq $false | Measure-Object).Count
            [PSCustomObject]@{
                "Project Key"  = $_.Name
                "Project Name" = $one.ProjectName
                "Actions OK"   = $ok
                "Actions KO"   = $ko
            }
        } | Sort-Object "Project Key"
}

# ----------------------------
# Export
# ----------------------------
$exportDir = Split-Path $exportFoundPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export found   : $exportFoundPath" -ForegroundColor Cyan
$resultFound | Export-Csv -Path $exportFoundPath -NoTypeInformation -Encoding UTF8

Write-Host "Export actions : $exportActionsPath" -ForegroundColor Cyan
$resultActions | Export-Csv -Path $exportActionsPath -NoTypeInformation -Encoding UTF8

Write-Host "Export summary : $exportSummaryPath" -ForegroundColor Cyan
$summary | Export-Csv -Path $exportSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Found   : $exportFoundPath"
Write-Host " - Actions : $exportActionsPath"
Write-Host " - Summary : $exportSummaryPath"
Write-Host " - Creds   : $credPath"