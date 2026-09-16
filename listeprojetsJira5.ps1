<#
Export Projets Jira + Schemes + Roles (Jira Cloud)
- Stockage local chiffré (DPAPI) de l'email+token (pas de ressaisie)
- Project Lead corrigé (expand=lead + fallback GET /project/{key})
- TLS 1.2 + Proxy système Windows (PAC/WPAD/McAfee)
- Exporte 2 CSV:
  - C:\Temp\liste_projets_jira.csv
  - C:\Temp\liste_roles_projets_jira.csv

Usage:
  .\listeprojetsJira4.ps1
  .\listeprojetsJira4.ps1 -ResetCreds
#>

param(
    [switch]$ResetCreds
)

# ----------------------------
# Runtime
# ----------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Jira Cloud => TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Proxy système Windows + creds Windows (utile derrière McAfee)
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

$exportProjectsPath = "C:\Temp\liste_projets_jira.csv"
$exportRolesPath    = "C:\Temp\liste_roles_projets_jira.csv"

# Fichier de credential (DPAPI user scope)
$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }

    if (-not $reset -and (Test-Path $credPath)) {
        try {
            return Import-Clixml -Path $credPath
        } catch {
            Write-Warning "Impossible de relire le fichier credential, nouvelle saisie requise. Détail: $_"
        }
    }

    $email = Read-Host "Email Atlassian (ex: prenom.nom@domaine.fr)"
    # Nettoyage si collé au format Markdown : [mail](mailto:mail)
    if ($email -match '^\[(.+?)\]\(mailto:(.+?)\)$') { $email = $Matches[2] }
    $email = ($email -replace '^mailto:', '').Trim()

    $secureToken = Read-Host "API Token Atlassian (saisie masquée)" -AsSecureString

    $cred = New-Object System.Management.Automation.PSCredential($email, $secureToken)
    $cred | Export-Clixml -Path $credPath

    Write-Host "Identifiants sauvegardés dans: $credPath" -ForegroundColor Green
    return $cred
}

$jiraCred = Get-JiraCredential -credPath $credPath -reset:$ResetCreds

# Conversion SecureString -> string (uniquement pour construire Basic Auth)
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
# Dernière activité (GET /search/jql)
# ----------------------------
function Get-ProjectLastWorkUpdate($siteUrl, $headers, $projectKey) {
    $jql = "project = $projectKey ORDER BY updated DESC"
    $encodedJql = [System.Uri]::EscapeDataString($jql)
    $url = "$siteUrl/rest/api/3/search/jql?jql=$encodedJql&maxResults=1&fields=updated"

    $resp = Invoke-JiraRest "GET" $url $headers
    if ($resp.issues -and $resp.issues.Count -gt 0) {
        return $resp.issues[0].fields.updated
    }
    return $null
}

# ----------------------------
# Project Lead (fallback)
# ----------------------------
function Get-ProjectLeadFallback($siteUrl, $headers, $projectKey) {
    $p = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey" $headers
    if ($p.lead) {
        return @{
            DisplayName = $p.lead.displayName
            AccountId   = $p.lead.accountId
        }
    }
    return @{ DisplayName = $null; AccountId = $null }
}

# ----------------------------
# Schemes par projet (endpoints officiels)
# ----------------------------
function Get-ProjectSchemes($siteUrl, $headers, $projectId, $projectKey) {
    $result = @{
        SchemeWorkType      = $null  # Issue Type Scheme
        SchemeWorkflow      = $null  # Workflow Scheme
        SchemeScreens       = $null  # Issue Type Screen Scheme
        SchemeFields        = $null  # Field Configuration Scheme
        SchemeNotification  = $null  # Notification Scheme
        SchemePermissions   = $null  # Permission Scheme
    }

    try {
        $its = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/issuetypescheme/project?projectId=$projectId" $headers
        if ($its.values -and $its.values.Count -gt 0) { $result.SchemeWorkType = $its.values[0].issueTypeScheme.name }
    } catch { }

    try {
        $wfs = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/workflowscheme/project?projectId=$projectId" $headers
        if ($wfs.values -and $wfs.values.Count -gt 0) { $result.SchemeWorkflow = $wfs.values[0].workflowScheme.name }
    } catch { }

    try {
        $itss = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/issuetypescreenscheme/project?projectId=$projectId" $headers
        if ($itss.values -and $itss.values.Count -gt 0) { $result.SchemeScreens = $itss.values[0].issueTypeScreenScheme.name }
    } catch { }

    try {
        $fcs = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/fieldconfigurationscheme/project?projectId=$projectId" $headers
        if ($fcs.values -and $fcs.values.Count -gt 0) { $result.SchemeFields = $fcs.values[0].fieldConfigurationScheme.name }
    } catch { }

    try {
        $ns = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey/notificationscheme" $headers
        if ($ns.name) { $result.SchemeNotification = $ns.name }
        elseif ($ns.notificationScheme -and $ns.notificationScheme.name) { $result.SchemeNotification = $ns.notificationScheme.name }
    } catch { }

    try {
        $ps = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey/permissionscheme" $headers
        if ($ps.name) { $result.SchemePermissions = $ps.name }
        elseif ($ps.permissionScheme -and $ps.permissionScheme.name) { $result.SchemePermissions = $ps.permissionScheme.name }
    } catch { }

    return $result
}

# ----------------------------
# Récupération projets (status=live) + expand=lead
# ----------------------------
Write-Host "Recuperation des projets (status=live)..." -ForegroundColor Cyan
$allProjects = @()
$startAt = 0
$maxResults = 50

while ($true) {
    $url = "$siteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$maxResults&status=live&expand=lead"
    $resp = Invoke-JiraRest "GET" $url $headers

    if ($resp.values) { $allProjects += $resp.values }
    $startAt += ($resp.values | Measure-Object).Count

    if ($startAt -ge $resp.total) { break }
}

Write-Host ("Nombre de projets : " + $allProjects.Count) -ForegroundColor Green

# ----------------------------
# Dataset Projets
# ----------------------------
Write-Host "Construction export Projets..." -ForegroundColor Cyan
$resultProjects = @()

foreach ($p in $allProjects) {
    $projectId   = $p.id
    $projectKey  = $p.key
    $projectName = $p.name

    if ([string]::IsNullOrWhiteSpace($projectKey) -or [string]::IsNullOrWhiteSpace($projectName)) {
        Write-Warning "Projet ignore (key/name vide)."
        continue
    }

    Write-Host "Projet $projectKey - $projectName" -ForegroundColor Yellow

    # Lead depuis search?expand=lead
    $leadDisplay = $null
    $leadAccount = $null
    if ($p.lead) {
        $leadDisplay = $p.lead.displayName
        if ($p.lead.accountId) { $leadAccount = $p.lead.accountId }
    }

    # Fallback si absent
    if ([string]::IsNullOrWhiteSpace($leadDisplay)) {
        try {
            $leadInfo = Get-ProjectLeadFallback $siteUrl $headers $projectKey
            $leadDisplay = $leadInfo.DisplayName
            $leadAccount = $leadInfo.AccountId
        } catch {
            Write-Warning "Lead non récupéré pour $projectKey : $_"
        }
    }

    $cat = $null
    if ($p.projectCategory) { $cat = $p.projectCategory.name }

    $lastUpdate = $null
    try { $lastUpdate = Get-ProjectLastWorkUpdate $siteUrl $headers $projectKey }
    catch { Write-Warning "Last update non recupere pour $projectKey : $_" }

    $schemes = $null
    try { $schemes = Get-ProjectSchemes $siteUrl $headers $projectId $projectKey }
    catch { Write-Warning "Schemes non recupere pour $projectKey : $_" }

    $resultProjects += [PSCustomObject]@{
        "Project Name"           = $projectName
        "Project Key"            = $projectKey
        "Project Lead"           = $leadDisplay
        "Project Lead AccountId" = $leadAccount
        "Project Category"       = $cat
        "Last Work Update"       = $lastUpdate
        "Scheme Work Type"       = $schemes.SchemeWorkType
        "Scheme Workflow"        = $schemes.SchemeWorkflow
        "Scheme Screens"         = $schemes.SchemeScreens
        "Scheme Fields"          = $schemes.SchemeFields
        "Scheme Notification"    = $schemes.SchemeNotification
        "Scheme Permissions"     = $schemes.SchemePermissions
    }
}

# ----------------------------
# Dataset Rôles + acteurs
# ----------------------------
Write-Host "Construction export Roles/Acteurs..." -ForegroundColor Cyan
$resultRoles = @()

foreach ($p in $allProjects) {
    $projectKey  = $p.key
    $projectName = $p.name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    Write-Host "Roles $projectKey - $projectName" -ForegroundColor Yellow

    $rolesResp = $null
    try {
        $rolesResp = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey/role" $headers
    } catch {
        Write-Warning "Roles non recupere pour $projectKey : $_"
        continue
    }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl = $prop.Value

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRest "GET" $roleUrl $headers
        } catch {
            Write-Warning "Detail role non recupere ($($prop.Name)) pour $projectKey : $_"
            continue
        }

        foreach ($actor in $roleDetail.actors) {
            $resultRoles += [PSCustomObject]@{
                "Project Key"   = $projectKey
                "Project Name"  = $projectName
                "Role Name"     = $roleDetail.name
                "Actor Type"    = $actor.type
                "Actor Name"    = $actor.displayName
                "Actor Account" = $actor.name
            }
        }
    }
}

# ----------------------------
# Export CSV
# ----------------------------
if (-not (Test-Path (Split-Path $exportProjectsPath))) {
    New-Item -ItemType Directory -Path (Split-Path $exportProjectsPath) | Out-Null
}

Write-Host "Export CSV projets: $exportProjectsPath" -ForegroundColor Cyan
$resultProjects | Export-Csv -Path $exportProjectsPath -NoTypeInformation -Encoding UTF8

Write-Host "Export CSV roles: $exportRolesPath" -ForegroundColor Cyan
$resultRoles | Export-Csv -Path $exportRolesPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Projets: $exportProjectsPath"
Write-Host " - Roles  : $exportRolesPath"
Write-Host " - Creds  : $credPath"