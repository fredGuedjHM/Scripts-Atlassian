<#
Retire les utilisateurs INACTIFS des groupes Jira dont le nom contient "DSIM"
- TLS 1.2 + Proxy système Windows (PAC/WPAD/McAfee)
- Stockage local chiffré (DPAPI) email+token
- DryRun par défaut (ne supprime rien) -> utiliser -Execute pour appliquer

Exports:
  * C:\Temp\dsim_inactive_memberships_found.csv
  * C:\Temp\dsim_inactive_memberships_actions.csv
  * C:\Temp\dsim_inactive_memberships_summary.csv

Usage:
  .\removeInactiveUsersFromDsimGroups.ps1
  .\removeInactiveUsersFromDsimGroups.ps1 -Execute
  .\removeInactiveUsersFromDsimGroups.ps1 -Filter "DSIM" -Execute
  .\removeInactiveUsersFromDsimGroups.ps1 -ResetCreds
#>

param(
    [string]$Filter = "DSIM",
    [switch]$Execute,
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
            # Retry simple sur 429
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

$exportFoundPath   = "C:\Temp\dsim_inactive_memberships_found.csv"
$exportActionsPath = "C:\Temp\dsim_inactive_memberships_actions.csv"
$exportSummaryPath = "C:\Temp\dsim_inactive_memberships_summary.csv"

# DPAPI credential file (user scope)
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
Write-Host ("Mode: " + $(if ($Execute) { "EXECUTE (suppression réelle)" } else { "DRYRUN (aucune suppression)" })) -ForegroundColor Magenta

# ----------------------------
# 1) Recherche groupes DSIM
# ----------------------------
Write-Host "Recherche des groupes contenant: '$Filter' ..." -ForegroundColor Cyan
$encoded = [System.Uri]::EscapeDataString($Filter)
$groupsUrl = "$siteUrl/rest/api/3/groups/picker?query=$encoded&maxResults=1000&caseInsensitive=true"
$groupsResp = Invoke-JiraRestWithRetry "GET" $groupsUrl $headers

$groups = @()
if ($groupsResp.groups) { $groups = $groupsResp.groups }
$groups = $groups | Where-Object { $_.name -and ($_.name.ToUpper().Contains($Filter.ToUpper())) } | Sort-Object name

Write-Host ("Groupes trouvés: " + $groups.Count) -ForegroundColor Green
if ($groups.Count -eq 0) { throw "Aucun groupe ne correspond au filtre '$Filter'." }

# ----------------------------
# 2) Membres inactifs trouvés (1 ligne = user x groupe)
# ----------------------------
Write-Host "Récupération des membres INACTIFS (active=false)..." -ForegroundColor Cyan

$found = @()

foreach ($g in $groups) {
    $gName = $g.name
    $gId   = $g.groupId

    Write-Host "Groupe: $gName" -ForegroundColor Yellow

    $startAt = 0
    $maxResults = 50
    while ($true) {
        $url = $null
        if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $url = "$siteUrl/rest/api/3/group/member?groupId=$([System.Uri]::EscapeDataString($gId))&includeInactiveUsers=true&startAt=$startAt&maxResults=$maxResults"
        } else {
            $url = "$siteUrl/rest/api/3/group/member?groupname=$([System.Uri]::EscapeDataString($gName))&includeInactiveUsers=true&startAt=$startAt&maxResults=$maxResults"
        }

        $page = Invoke-JiraRestWithRetry "GET" $url $headers
        $values = @()
        if ($page.values) { $values = $page.values }

        foreach ($u in $values) {
            if ($u.active -eq $false) {
                $mail = $null
                if ($u.PSObject.Properties.Name -contains "emailAddress") { $mail = $u.emailAddress }

                $found += [PSCustomObject]@{
                    "Group Name"   = $gName
                    "Group Id"     = $gId
                    "AccountId"    = $u.accountId
                    "DisplayName"  = $u.displayName
                    "EmailAddress" = $mail
                    "Active"       = $u.active
                    "AccountType"  = $u.accountType
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

Write-Host ("Memberships inactifs trouvés: " + $found.Count) -ForegroundColor Green

# ----------------------------
# 3) Suppression (ou simulation) des memberships
# ----------------------------
Write-Host "Traitement des suppressions..." -ForegroundColor Cyan

$actions = @()
$ts = (Get-Date).ToString("s")

foreach ($m in $found) {
    $gName = $m."Group Name"
    $gId   = $m."Group Id"
    $aid   = $m.AccountId
    $disp  = $m.DisplayName

    if ([string]::IsNullOrWhiteSpace($aid)) { continue }

    if (-not $Execute) {
        $actions += [PSCustomObject]@{
            Timestamp   = $ts
            Mode        = "DRYRUN"
            Action      = "WOULD_REMOVE_USER_FROM_GROUP"
            Success     = $true
            GroupName   = $gName
            GroupId     = $gId
            AccountId   = $aid
            DisplayName = $disp
            Details     = "DELETE /rest/api/3/group/user?groupId=<id>&accountId=<id>"
        }
        continue
    }

    # EXECUTE
    try {
        $delUrl = $null
        if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $delUrl = "$siteUrl/rest/api/3/group/user?groupId=$([System.Uri]::EscapeDataString($gId))&accountId=$([System.Uri]::EscapeDataString($aid))"
        } else {
            $delUrl = "$siteUrl/rest/api/3/group/user?groupname=$([System.Uri]::EscapeDataString($gName))&accountId=$([System.Uri]::EscapeDataString($aid))"
        }

        Invoke-JiraRestWithRetry "DELETE" $delUrl $headers | Out-Null

        $actions += [PSCustomObject]@{
            Timestamp   = $ts
            Mode        = "EXECUTE"
            Action      = "REMOVE_USER_FROM_GROUP"
            Success     = $true
            GroupName   = $gName
            GroupId     = $gId
            AccountId   = $aid
            DisplayName = $disp
            Details     = $delUrl
        }
    } catch {
        $actions += [PSCustomObject]@{
            Timestamp   = $ts
            Mode        = "EXECUTE"
            Action      = "REMOVE_USER_FROM_GROUP_FAILED"
            Success     = $false
            GroupName   = $gName
            GroupId     = $gId
            AccountId   = $aid
            DisplayName = $disp
            Details     = "$_"
        }
    }
}

# ----------------------------
# 4) Summary par user
# ----------------------------
$summary = @()
if ($actions.Count -gt 0) {
    $summary = $actions |
        Where-Object { $_.Action -in @("WOULD_REMOVE_USER_FROM_GROUP","REMOVE_USER_FROM_GROUP","REMOVE_USER_FROM_GROUP_FAILED") } |
        Group-Object AccountId |
        ForEach-Object {
            $one = $_.Group | Select-Object -First 1
            $groups = ($_.Group | Select-Object -ExpandProperty GroupName -Unique | Sort-Object) -join ";"
            $ok = ($_.Group | Where-Object Action -eq "REMOVE_USER_FROM_GROUP" | Measure-Object).Count
            $ko = ($_.Group | Where-Object Action -eq "REMOVE_USER_FROM_GROUP_FAILED" | Measure-Object).Count
            $would = ($_.Group | Where-Object Action -eq "WOULD_REMOVE_USER_FROM_GROUP" | Measure-Object).Count

            [PSCustomObject]@{
                AccountId            = $one.AccountId
                DisplayName          = $one.DisplayName
                Groups               = $groups
                WouldRemoveCount     = $would
                RemoveSuccessCount   = $ok
                RemoveFailedCount    = $ko
            }
        } | Sort-Object DisplayName
}

# ----------------------------
# Export CSV
# ----------------------------
$exportDir = Split-Path $exportFoundPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export found   : $exportFoundPath" -ForegroundColor Cyan
$found | Export-Csv -Path $exportFoundPath -NoTypeInformation -Encoding UTF8

Write-Host "Export actions : $exportActionsPath" -ForegroundColor Cyan
$actions | Export-Csv -Path $exportActionsPath -NoTypeInformation -Encoding UTF8

Write-Host "Export summary : $exportSummaryPath" -ForegroundColor Cyan
$summary | Export-Csv -Path $exportSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Found   : $exportFoundPath"
Write-Host " - Actions : $exportActionsPath"
Write-Host " - Summary : $exportSummaryPath"
Write-Host " - Creds   : $credPath"