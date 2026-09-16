<#
Export membres des groupes Jira dont le nom contient "DSIM" (Jira Cloud)
- TLS 1.2 + Proxy système Windows (PAC/WPAD/McAfee)
- Stockage local chiffré (DPAPI) de l'email+token (pas de ressaisie)
- Exporte 3 CSV:
    * C:\Temp\groupes_DSIM.csv
    * C:\Temp\membres_groupes_DSIM_detail.csv
    * C:\Temp\membres_groupes_DSIM_uniques.csv

Usage:
  .\listeGroupesDsimMembres.ps1
  .\listeGroupesDsimMembres.ps1 -Filter "DSIM"
  .\listeGroupesDsimMembres.ps1 -ResetCreds
#>

param(
    [string]$Filter = "DSIM",
    [switch]$IncludeInactiveUsers,
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
        $msg = $_.Exception.Message
        $details = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = $_.ErrorDetails.Message }
        if ($_.Exception.InnerException) { $msg = "$msg | Inner: $($_.Exception.InnerException.Message)" }

        if ($details) { throw "HTTP failed: $method $url => $msg | Details: $details" }
        throw "HTTP failed: $method $url => $msg"
    }
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportGroupsPath  = "C:\Temp\groupes_DSIM.csv"
$exportDetailPath  = "C:\Temp\membres_groupes_DSIM_detail.csv"
$exportUniquePath  = "C:\Temp\membres_groupes_DSIM_uniques.csv"

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
$me = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

# ----------------------------
# 1) Recherche des groupes contenant $Filter
# ----------------------------
Write-Host "Recherche des groupes dont le nom contient: '$Filter' ..." -ForegroundColor Cyan
$encoded = [System.Uri]::EscapeDataString($Filter)

# Find groups (groups picker)
# NB: le maxResults est plafonné côté Jira (jira.ajax.autocomplete.limit)
$groupsUrl = "$siteUrl/rest/api/3/groups/picker?query=$encoded&maxResults=1000&caseInsensitive=true"
$groupsResp = Invoke-JiraRest "GET" $groupsUrl $headers

$groups = @()
if ($groupsResp.groups) { $groups = $groupsResp.groups }

# Sécurité : filtrage local "contains" au cas où
$groups = $groups | Where-Object { $_.name -and ($_.name.ToUpper().Contains($Filter.ToUpper())) }

Write-Host ("Groupes trouvés: " + $groups.Count) -ForegroundColor Green

# Dataset groupes
$resultGroups = $groups | Sort-Object name | ForEach-Object {
    [PSCustomObject]@{
        "Group Name" = $_.name
        "Group Id"   = $_.groupId
    }
}

# ----------------------------
# 2) Pour chaque groupe => liste paginée des membres
# ----------------------------
Write-Host "Récupération des membres par groupe..." -ForegroundColor Cyan

$resultDetail = @()

foreach ($g in $groups) {
    $gName = $g.name
    $gId   = $g.groupId

    if ([string]::IsNullOrWhiteSpace($gId) -and [string]::IsNullOrWhiteSpace($gName)) { continue }

    Write-Host "Groupe: $gName" -ForegroundColor Yellow

    $startAt = 0
    $maxResults = 50

    while ($true) {
        $incInactive = $(if ($IncludeInactiveUsers) { "true" } else { "false" })

        # On privilégie groupId (plus stable), sinon groupname
        if (-not [string]::IsNullOrWhiteSpace($gId)) {
            $url = "$siteUrl/rest/api/3/group/member?groupId=$([System.Uri]::EscapeDataString($gId))&includeInactiveUsers=$incInactive&startAt=$startAt&maxResults=$maxResults"
        } else {
            $url = "$siteUrl/rest/api/3/group/member?groupname=$([System.Uri]::EscapeDataString($gName))&includeInactiveUsers=$incInactive&startAt=$startAt&maxResults=$maxResults"
        }

        $page = $null
        try {
            $page = Invoke-JiraRest "GET" $url $headers
        } catch {
            Write-Warning "Impossible de lire les membres du groupe '$gName' : $_"
            break
        }

        $values = @()
        if ($page.values) { $values = $page.values }

        foreach ($u in $values) {
            # emailAddress peut être null/absent selon privacy + permissions
            $mail = $null
            if ($u.PSObject.Properties.Name -contains "emailAddress") { $mail = $u.emailAddress }

            $resultDetail += [PSCustomObject]@{
                "Group Name"    = $gName
                "Group Id"      = $gId
                "AccountId"     = $u.accountId
                "DisplayName"   = $u.displayName
                "EmailAddress"  = $mail
                "Active"        = $u.active
                "AccountType"   = $u.accountType
            }
        }

        $returned = ($values | Measure-Object).Count
        $startAt += $returned

        # arrêt pagination
        if ($returned -eq 0) { break }
        if ($page.isLast -eq $true) { break }
        if ($page.total -and $startAt -ge $page.total) { break }
    }
}

Write-Host ("Lignes détail membres: " + $resultDetail.Count) -ForegroundColor Green

# ----------------------------
# 3) Liste unique des personnes + groupes associés
# ----------------------------
$resultUnique = @()
if ($resultDetail.Count -gt 0) {
    $resultUnique = $resultDetail |
        Group-Object AccountId |
        ForEach-Object {
            $one = $_.Group | Select-Object -First 1
            $groupsList = ($_.Group | Select-Object -ExpandProperty "Group Name" | Sort-Object -Unique) -join ";"

            [PSCustomObject]@{
                "AccountId"    = $one.AccountId
                "DisplayName"  = $one.DisplayName
                "EmailAddress" = $one.EmailAddress
                "Active"       = $one.Active
                "Groups"       = $groupsList
            }
        } |
        Sort-Object DisplayName
}

# ----------------------------
# Export CSV
# ----------------------------
$exportDir = Split-Path $exportGroupsPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export CSV groupes : $exportGroupsPath" -ForegroundColor Cyan
$resultGroups | Export-Csv -Path $exportGroupsPath -NoTypeInformation -Encoding UTF8

Write-Host "Export CSV détail membres : $exportDetailPath" -ForegroundColor Cyan
$resultDetail | Export-Csv -Path $exportDetailPath -NoTypeInformation -Encoding UTF8

Write-Host "Export CSV membres uniques : $exportUniquePath" -ForegroundColor Cyan
$resultUnique | Export-Csv -Path $exportUniquePath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Groupes        : $exportGroupsPath"
Write-Host " - Membres détail : $exportDetailPath"
Write-Host " - Membres uniques: $exportUniquePath"
Write-Host " - Creds          : $credPath"